using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Office365ManagementAPIWebHook.Model;
/*
 * Create Azure Function in Visual Studio 2022. Add the below code and run .\ngrok.exe http 7141 in local machine to test
 * This sample just print the log when run in local machine
 * In real production code, we need to save the notifications to the storage for further processing
 * Not verify the certificate from manage.office.com
*/
namespace Office365ManagementAPIWebHook
{
    public class O365ManagementAPIHttpFunction
    {
        private readonly ILogger<O365ManagementAPIHttpFunction> _logger;
        private readonly string webhookValidationcodString ="Webhook-Validationcode";
        private readonly string webhookAuthidString = "Webhook-Authid";
        private readonly string tenantAuthId = "ZiZhuOffice365ManagementAPINotification20240220";

        public O365ManagementAPIHttpFunction(ILogger<O365ManagementAPIHttpFunction> logger)
        {
            _logger = logger;
        }

        [Function("O365ManagementAPIHttpFunction")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var reqHeaders = req.Headers;
            if(reqHeaders.ContainsKey(webhookValidationcodString) && reqHeaders.ContainsKey(webhookAuthidString) &&
                string.Equals(reqHeaders[webhookAuthidString], tenantAuthId, StringComparison.InvariantCultureIgnoreCase))
            {
                _logger.LogInformation("C# HTTP trigger function passed the webhook validation.");
                return new OkResult();
            }
            
            if (!reqHeaders.ContainsKey(webhookValidationcodString) && reqHeaders.ContainsKey(webhookAuthidString) &&
                string.Equals(reqHeaders[webhookAuthidString], tenantAuthId, StringComparison.InvariantCultureIgnoreCase))
            {
                _logger.LogInformation("C# HTTP trigger function received the notifications.");
                var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var data = JsonConvert.DeserializeObject<List<ContentNotification>>(requestBody);
                data?.ForEach(c => _logger.LogInformation(c.contentUri.ToString()));
            }
            return new OkResult();
        }
    }
}
