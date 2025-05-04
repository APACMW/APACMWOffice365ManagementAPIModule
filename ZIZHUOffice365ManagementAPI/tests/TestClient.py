import datetime
from ZIZHUOffice365ManagementAPI import Office365ManagementAPIServiceClient, ContentType
from cryptography.hazmat.primitives.serialization import (
    pkcs12,
    Encoding,
    PrivateFormat,
    NoEncryption,
)
from cryptography.hazmat.primitives import hashes
'''
# 1 use the client secret to authenticate
tenant_id = 'cff343b2-f0ff-416a-802b-28595997daa2'
client_id = '9b0547c4-28b1-466d-a80e-677c6dc42d42'
loginHint = 'freeman@vjqg8.onmicrosoft.com'
client = Office365ManagementAPIServiceClient(
    clientId=client_id,
    tenantID=tenant_id,
    loginHint=loginHint
)
tenant_id = 'cff343b2-f0ff-416a-802b-28595997daa2'
client_id = 'bc4db1db-b705-434a-91ff-145aa94185c8'
client_secret = ''
client = Office365ManagementAPIServiceClient(
    clientId=client_id,
    tenantID=tenant_id,
    clientsecret=client_secret
)

# 2 use the certificate to authenticate
tenant_id = "cff343b2-f0ff-416a-802b-28595997daa2"
client_id = "bc4db1db-b705-434a-91ff-145aa94185c8"
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://manage.office.com/.default"]
certificate_path = "c:/temp/O365AppCertificate.pfx"
certificate_password = "1234qwer"

# Load the certificate
with open(certificate_path, "rb") as cert_file:
    pfx_data = cert_file.read()

private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
    pfx_data, certificate_password.encode() if certificate_password else None
)

# Convert the private key to PEM format
private_key_pem = private_key.private_bytes(
    encoding=Encoding.PEM,
    format=PrivateFormat.PKCS8,
    encryption_algorithm=NoEncryption(),
).decode("utf-8")

# Get the thumbprint of the certificate
thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
client = Office365ManagementAPIServiceClient(
    clientId=client_id,
    tenantID=tenant_id,
    certificateprivatekey=private_key_pem,
    certificatethumbprint=thumbprint
)

# 3 interactive authentication
tenant_id = 'cff343b2-f0ff-416a-802b-28595997daa2'
client_id = '9b0547c4-28b1-466d-a80e-677c6dc42d42'
loginHint = 'freeman@vjqg8.onmicrosoft.com'
client = Office365ManagementAPIServiceClient(
    clientId=client_id,
    tenantID=tenant_id,
    loginHint=loginHint
)
'''
tenant_id = "cff343b2-f0ff-416a-802b-28595997daa2"
client_id = "bc4db1db-b705-434a-91ff-145aa94185c8"
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://manage.office.com/.default"]
certificate_path = "c:/temp/O365AppCertificate.pfx"
certificate_password = "1234qwer"

# Load the certificate
with open(certificate_path, "rb") as cert_file:
    pfx_data = cert_file.read()

private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
    pfx_data, certificate_password.encode() if certificate_password else None
)

# Convert the private key to PEM format
private_key_pem = private_key.private_bytes(
    encoding=Encoding.PEM,
    format=PrivateFormat.PKCS8,
    encryption_algorithm=NoEncryption(),
).decode("utf-8")

# Get the thumbprint of the certificate
thumbprint = certificate.fingerprint(hashes.SHA1()).hex()
client = Office365ManagementAPIServiceClient(
    clientId=client_id,
    tenantID=tenant_id,
    certificateprivatekey=private_key_pem,
    certificatethumbprint=thumbprint
)

client.connect()
client.stop_subscriptions()
contenttypes = [ContentType.AuditExchange, ContentType.AuditSharePoint,
                ContentType.AuditAzureActiveDirectory, ContentType.AuditGeneral, ContentType.DLPAll]
for contenttype in contenttypes:
    client.start_subscription(contenttype)
print(client.get_currentsubscriptions())
startTime = datetime.datetime(2025, 4, 28, 0, 0, 0)
endTime = datetime.datetime(2025, 4, 29, 0, 0, 0)

for contenttype in contenttypes:
    blobs = client.get_availablecontent(startTime, endTime, contenttype)
    client.receive_content(blobs)
    print(blobs)
print(client.receive_resourcefriendlynames())
client.disconnect()
