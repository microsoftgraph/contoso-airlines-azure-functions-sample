using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CreateFlightTeam
{
    [StorageAccount("AzureWebJobsStorage")]
    public static class ConnectGraphNotifications
    {
        [FunctionName("ConnectGraphNotifications")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // POST comes from mobile app and should have a bearer token
            var authHeader = req.Headers["Authorization"];

            try
            {
                await Graph.AuthProvider.GetTokenOnBehalfOfAsync(authHeader, log);
                // Return 202
                return new AcceptedResult();
            }
            catch (MsalException ex)
            {
                log.LogError(ex.Message);
                return new BadRequestResult();
            }
        }
    }
}