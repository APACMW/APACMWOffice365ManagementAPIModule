import datetime
from enum import Enum
from dataclasses import dataclass
from typing import List
from cryptography import x509
from msal import PublicClientApplication, ConfidentialClientApplication
import logging
import requests
import time
import math
import msal
import json


class Office365SubscriptionPlanType(Enum):
    Enterpriseplan = 1
    GCCGovernmentPlan = 2
    GCCHighGovernmentPlan = 3
    DoDGovernmentPlan = 4
    GallatinPlan = 5


class ContentType(Enum):
    AuditAzureActiveDirectory = 1
    AuditExchange = 2
    AuditSharePoint = 3
    AuditGeneral = 4
    DLPAll = 5


@dataclass
class Blob:
    contentUri: str
    contentId: str
    contentType: str
    contentCreated: datetime
    contentExpiration: datetime
    auditRecords: list[str]


@dataclass
class WebHook:
    authId: str
    address: str
    expiration: str
    status: str


@dataclass
class Subscription:
    contentType: str
    status: str
    webhook: WebHook


@dataclass
class Notification:
    contentType: str
    contentId: str
    contentUri: str
    notificationStatus: str
    contentCreated: datetime
    notificationSent: datetime
    contentExpiration: datetime


@dataclass
class Resourcefriendlyname:
    id: str
    name: str


class Office365ManagementAPIServiceClient:
    max_retries: int = 4
    sleep_seconds: int = 3

    def __init__(
        self,
        tenantID: str,
        clientId: str,
        loginHint: str = None,
        clientsecret: str = None,
        certificateprivatekey: str = None,
        certificatethumbprint: str = None,
        is_disposed: bool = False,
        office365SubscriptionPlanType: Office365SubscriptionPlanType = Office365SubscriptionPlanType.Enterpriseplan,
        log_level: int = logging.WARNING,
    ):
        self.is_disposed = is_disposed
        self.tenantID = tenantID
        self.clientId = clientId
        self.loginHint = loginHint
        self.clientsecret = clientsecret
        self.certificateprivatekey = certificateprivatekey
        self.certificatethumbprint = certificatethumbprint
        self.office365SubscriptionPlanType = office365SubscriptionPlanType
        self.authResult = None
        self.auth_app = None
        self.token_cache = msal.SerializableTokenCache()
        self.httpErrorResponse = None
        self.azureCloudInstance = None
        self.root = None
        self.scope = None
        current_datetime = datetime.datetime.now()
        formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"./ZizhuOffice365ManagementAPI_{formatted_datetime}.log"
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(log_level)
        console_handler = logging.StreamHandler()
        file_handler = logging.FileHandler(filename)
        console_handler.setLevel(log_level)
        file_handler.setLevel(log_level)
        logger_format = logging.Formatter(
            "%(name)s - %(asctime)s[%(levelname)s]%(funcName)s: %(message)s"
        )
        console_handler.setFormatter(logger_format)
        file_handler.setFormatter(logger_format)
        self.logger.addHandler(console_handler)
        self.logger.addHandler(file_handler)
        self._setroot(office365SubscriptionPlanType)

    def dispose(self):
        if not self.is_disposed:
            if self.token_cache != None:
                del self.token_cache
            if self.authResult != None:
                del self.authResult
            self.tenantID = None
            self.clientId = None
            self.clientsecret = None
            self.certificateprivatekey = None
            self.loginHint = None
            self.certificateprivatekey = None
            self.certificatethumbprint = None
            self.office365SubscriptionPlanType = None
            self.auth_app = None
            self.root = None
            self.scope = None
            self.azureCloudInstance = None
            self.httpErrorResponse = None
            self.is_disposed = True

    def __del__(self):
        self.dispose()

    def _setroot(self, office365SubscriptionPlanType: Office365SubscriptionPlanType):
        match office365SubscriptionPlanType:
            case Office365SubscriptionPlanType.Enterpriseplan:
                self.root = "https://manage.office.com"
                self.azureCloudInstance = "AzurePublic"
            case Office365SubscriptionPlanType.GCCHighGovernmentPlan:
                self.root = "https://manage.office365.us"
                self.azureCloudInstance = "AzurePublic"
            case Office365SubscriptionPlanType.china:
                self.root = "https://manage.office365.cn"
                self.azureCloudInstance = "AzureChina"
            case _:
                raise ValueError("Invalid Office365SubscriptionPlanType")
        self.scope = [f"{self.root}/.default"]
        self.logger.info(
            f"Root: {self.root}; AzureCloudInstance: {self.azureCloudInstance}; Scope: {self.scope}"
        )

    def disconnect(self):
        self.dispose()
        self.logger.debug("Disconnected from Office 365 Management API")
        return

    def connect(self):
        if not self._get_oauth_toekn():
            self.logger.error("Failed to get access token")
            self.logger.debug(
                f"Error: {self.httpErrorResponse.status_code} {self.httpErrorResponse.text}"
            )
            raise RuntimeError(
                "Can not get the access token. Enable the debug log to investigate."
            )
        else:
            self.logger.info("Access token acquired")
        return

    def _get_oauth_toekn(self) -> bool:
        result = False
        try:
            authority = f"https://login.microsoftonline.com/{self.tenantID}"
            self.authResult = None
            if self.loginHint != None:
                if self.auth_app == None:
                    self.auth_app = PublicClientApplication(
                        self.clientId, authority=authority, token_cache=self.token_cache
                    )

                accounts = self.auth_app.get_accounts()
                if accounts:
                    chosen_account = accounts[0]
                    self.authResult = self.auth_app.acquire_token_silent(
                        scopes=self.scope, account=chosen_account
                    )
                else:
                    self.authResult = self.auth_app.acquire_token_interactive(
                        scopes=self.scope, login_hint=self.loginHint
                    )
            elif self.clientsecret != None:
                if self.auth_app == None:
                    self.auth_app = ConfidentialClientApplication(
                        self.clientId,
                        authority=authority,
                        token_cache=self.token_cache,
                        client_credential=self.clientsecret,
                    )
                self.authResult = self.auth_app.acquire_token_silent(
                    scopes=self.scope, account=None
                )
                if self.authResult == None:
                    self.authResult = self.auth_app.acquire_token_for_client(
                        scopes=self.scope
                    )
            elif (
                self.certificateprivatekey != None
                and self.certificatethumbprint != None
            ):
                if self.auth_app == None:
                    self.auth_app = ConfidentialClientApplication(
                        self.clientId,
                        authority=authority,
                        token_cache=self.token_cache,
                        client_credential={
                            "private_key": self.certificateprivatekey,
                            "thumbprint": self.certificatethumbprint,
                        },
                    )
                self.authResult = self.auth_app.acquire_token_silent(
                    scopes=self.scope, account=None
                )
                if self.authResult == None:
                    self.authResult = self.auth_app.acquire_token_for_client(
                        scopes=self.scope
                    )
            else:
                self.logger.error("Invalid authentication method")
                raise ValueError("Invalid authentication method")

            if self.authResult != None and "access_token" in self.authResult:
                result = True
        except Exception as e:
            self.logger.error(f"Http error: {str(e)}")
        return result

    def _get_contentTypeEnum(self, contenttype_String: str) -> ContentType:
        """Get content type enum from string"""
        match contenttype_String:
            case "Audit.AzureActiveDirectory":
                return ContentType.AuditAzureActiveDirectory
            case "Audit.Exchange":
                return ContentType.AuditExchange
            case "Audit.SharePoint":
                return ContentType.AuditSharePoint
            case "Audit.General":
                return ContentType.AuditGeneral
            case "DLP.All":
                return ContentType.DLPAll
            case _:
                raise ValueError("Invalid content type string")

    def _get_contenttype_String(self, contentType: ContentType) -> str:
        """Get content type string"""
        match contentType:
            case ContentType.AuditAzureActiveDirectory:
                return "Audit.AzureActiveDirectory"
            case ContentType.AuditExchange:
                return "Audit.Exchange"
            case ContentType.AuditSharePoint:
                return "Audit.SharePoint"
            case ContentType.AuditGeneral:
                return "Audit.General"
            case ContentType.DLPAll:
                return "DLP.All"
            case _:
                raise ValueError("Invalid content type")

    def _invoke_o365managementapi_request(
        self, url, http_verb, request_body=None
    ) -> requests.Response:
        """Invoke the O365 Management API request and return the response"""
        http_response = None
        http_error_response = None
        retry_count = 0

        while retry_count <= Office365ManagementAPIServiceClient.max_retries:
            if retry_count > 1:
                sleep_seconds = math.pow(
                    Office365ManagementAPIServiceClient.sleep_seconds, retry_count
                )
                self.logger.info(
                    f"Retry Invoke-O365APIHttpRequest {http_verb} {url}: {retry_count} after {sleep_seconds} seconds"
                )
                time.sleep(sleep_seconds)

            headers = {
                "Authorization": f"{self.authResult['token_type']} {self.authResult['access_token']}",
                "Content-Type": "application/json",
            }
            try:
                if request_body:
                    http_response = requests.request(
                        http_verb, url, headers=headers, json=request_body
                    )
                else:
                    http_response: requests.Response = requests.request(
                        http_verb, url, headers=headers
                    )

                if http_response.status_code in [200, 204]:
                    self.logger.info(f"HTTP Response: {http_response.text}")
                    return http_response
                else:
                    self.httpErrorResponse = http_response
                    self.logger.info(
                        f"Http error: {http_response.status_code} Body: {http_response.text}"
                    )
            except requests.RequestException as e:
                self.logger.error(f"Http error: {str(e)}")
                self.httpErrorResponse = e.response
            except Exception as e:
                self.logger.error(f"Http error: {str(e)}")
            retry_count += 1

            if http_error_response and http_error_response.status_code in [401, 400]:
                self.logger.error(
                    f"Http error: {http_error_response.status_code} Body: {http_error_response.text}"
                )
                break
        return http_response

    def get_currentsubscriptions(self) -> List[Subscription]:
        """Get current subscriptions"""
        listSubscriptionURI = (
            f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/subscriptions/list"
        )
        http_response = self._invoke_o365managementapi_request(
            listSubscriptionURI, "GET"
        )
        if http_response.status_code != 200:
            self.logger.error(
                f"Failed to get subscriptions. Status code: {http_response.status_code}"
            )
            raise RuntimeError(
                f"Failed to get subscriptions. Status code: {http_response.status_code}"
            )
        subscriptions: List[Subscription] = []
        subscriptionObjects = json.loads(http_response.content)
        for subscription in subscriptionObjects:
            contentType = subscription["contentType"]
            status = subscription["status"]
            webhook = subscription["webhook"]
            webHookObject = None
            if webhook != None:
                webHookObject = WebHook(
                    authId=webhook["authId"],
                    address=webhook["address"],
                    expiration=webhook["expiration"],
                    status=webhook["status"],
                )
            subscriptionObject = Subscription(
                contentType=contentType,
                status=status,
                webhook=webHookObject,
            )
            subscriptions.append(subscriptionObject)
        return subscriptions

    def stop_subscription(self, contentType: ContentType) -> bool:
        """Stop a subscription"""
        contenttypeString = self._get_contenttype_String(contentType)
        subscriptionUrl = f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/subscriptions/stop?contentType={contenttypeString}"
        http_response = self._invoke_o365managementapi_request(subscriptionUrl, "Post")
        if http_response.status_code != 204:
            self.logger.error(
                f"Failed to get subscriptions. Status code: {http_response.status_code}"
            )
            raise RuntimeError(
                f"Failed to stop subscription. Status code: {http_response.status_code}"
            )
        return True

    def stop_subscriptions(self) -> bool:
        """Stop all subscriptions"""
        subscriptions = self.get_currentsubscriptions()
        for subscription in subscriptions:
            if subscription.status == "enabled":
                self.stop_subscription(
                    self._get_contentTypeEnum(subscription.contentType)
                )
        return True

    def start_subscription(self, contenttype: ContentType, webhook: str = None) -> bool:
        """Start a subscription"""
        contenttypeString = self._get_contenttype_String(contenttype)
        subscriptions = self.get_currentsubscriptions()
        for subscription in subscriptions:
            if (
                subscription.status == "enabled"
                and subscription.contentType == contenttypeString
            ):
                return True
        subscription_url = f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/subscriptions/start?contentType={contenttypeString}"
        if webhook != None:
            http_response = self._invoke_o365managementapi_request(
                subscription_url, "POST", webhook
            )
        else:
            http_response = self._invoke_o365managementapi_request(
                subscription_url, "POST"
            )
        if http_response.status_code != 200:
            self.logger.error(
                f"Failed to get subscriptions. Status code: {http_response.status_code}"
            )
            raise RuntimeError(
                f"Failed to start subscription. Status code: {http_response.status_code}"
            )
        return True

    def get_availablecontent(
        self, startTime: datetime, endTime: datetime, contentType: ContentType = None
    ) -> List[Blob]:
        startTimeStr = startTime.strftime("%Y-%m-%dT%H:%M:%S")
        endTimeStr = endTime.strftime("%Y-%m-%dT%H:%M:%S")
        contentTypeString = self._get_contenttype_String(contentType)
        subscriptions = self.get_currentsubscriptions()
        subscriptions = [sub for sub in subscriptions if sub.status == "enabled"]
        if contentType != None:
            subscriptions = [
                sub for sub in subscriptions if sub.contentType == contentTypeString
            ]
        if len(subscriptions) == 0:
            self.logger.info("No subscriptions found")
        contentTypes = [sub.contentType for sub in subscriptions]
        availableContent: List[Blob] = []
        for content_Type in contentTypes:
            contentUrl: str = (
                f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/subscriptions/content?contentType={content_Type}&startTime={startTimeStr}&endTime={endTimeStr}"
            )
            while contentUrl != None:
                http_response = self._invoke_o365managementapi_request(
                    contentUrl, "GET"
                )
                if http_response.status_code != 200:
                    self.logger.error(
                        f"Failed to get available content. Status code: {http_response.status_code}"
                    )
                    raise RuntimeError(
                        f"Failed to get available content. Status code: {http_response.status_code}"
                    )
                contentObjects = json.loads(http_response.content)
                for contentObject in contentObjects:
                    blobObject = Blob(
                        contentUri=contentObject["contentUri"],
                        contentId=contentObject["contentId"],
                        contentType=contentObject["contentType"],
                        contentCreated=contentObject["contentCreated"],
                        contentExpiration=contentObject["contentExpiration"],
                        auditRecords=None,
                    )
                    availableContent.append(blobObject)
                contentUrl = None
                contentUrl = http_response.headers.get("NextPageUri")
        return availableContent

    def receive_content(self, blobs: List[Blob]):
        """Receive content from blobs"""
        if blobs == None or len(blobs) == 0:
            return
        for blob in blobs:
            contentUrl = blob.contentUri
            http_response = self._invoke_o365managementapi_request(contentUrl, "GET")
            if http_response.status_code != 200:
                self.logger.error(
                    f"Failed to receive content. Status code: {http_response.status_code}"
                )
                raise RuntimeError(
                    f"Failed to receive content. Status code: {http_response.status_code}"
                )
            blob.auditRecords = json.loads(http_response.content)
        return blobs

    def receive_notifications(self):
        self.logger.error(
            "The content notifications are sent to the webhook. We can not implement this operation in Python Module."
        )
        return

    def get_notifications(
        self, startTime: datetime, endTime: datetime, contentType: ContentType = None
    ) -> List[Notification]:
        """Get notifications"""
        startTimeStr = startTime.strftime("%Y-%m-%dT%H:%M:%S")
        endTimeStr = endTime.strftime("%Y-%m-%dT%H:%M:%S")
        contentTypeString = self._get_contenttype_String(contentType)
        listNotificationsUrl: str = (
            f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/subscriptions/notifications?contentType={contentTypeString}&startTime={startTimeStr}&endTime={endTimeStr}"
        )
        http_response = self._invoke_o365managementapi_request(
            listNotificationsUrl, "Get"
        )
        if http_response.status_code != 200:
            raise RuntimeError(
                f"Failed to get notifications. Status code: {http_response.status_code}"
            )
        notificationObjects = json.loads(http_response.content)
        notifications: List[Notification] = []
        for notificationObject in notificationObjects:
            notification = Notification(
                contentType=notificationObject["contentType"],
                contentId=notificationObject["contentId"],
                contentUri=notificationObject["contentUri"],
                notificationStatus=notificationObject["notificationStatus"],
                contentCreated=notificationObject["contentCreated"],
                notificationSent=notificationObject["notificationSent"],
                contentExpiration=notificationObject["contentExpiration"],
            )
            notifications.append(notification)
        return notifications

    def receive_resourcefriendlynames(self) -> List[Resourcefriendlyname]:
        url = f"{self.root}/api/v1.0/{self.tenantID}/activity/feed/resources/dlpSensitiveTypes"
        http_response = self._invoke_o365managementapi_request(url, "Get")

        if http_response.status_code != 200:
            self.logger.error(
                f"Failed to get resource friendly names. Status code: {http_response.status_code}"
            )
            raise RuntimeError(
                f"Failed to get resource friendly names. Status code: {http_response.status_code}"
            )
        resourcefriendlyname_objects = json.loads(http_response.content)
        resourcefriendlynames: List[Resourcefriendlyname] = []
        for resourcefriendlyname_object in resourcefriendlyname_objects:
            resourcefriendlyname = Resourcefriendlyname(
                id=resourcefriendlyname_object["id"],
                name=resourcefriendlyname_object["name"],
            )
            resourcefriendlynames.append(resourcefriendlyname)
        return resourcefriendlynames
