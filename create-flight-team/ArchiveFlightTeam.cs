using create_flight_team.Graph;
using create_flight_team.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class ArchiveFlightTeam
    {
        private static TraceWriter logger = null;

        [FunctionName("ArchiveFlightTeam")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                log.Info($"Access token: {graphToken}");

                string requestBody = new StreamReader(req.Body).ReadToEnd();
                var request = JsonConvert.DeserializeObject<ArchiveTeamRequest>(requestBody);

                await ArchiveTeamAsync(graphToken, request);

                return new OkResult();
            }
            catch (AdalException ex)
            {
                log.Info($"Could not obtain Graph token: {ex.Message}");
                // Just return 401 if something went wrong
                // during token exchange
                return new UnauthorizedResult();
            }
            catch (Exception ex)
            {
                log.Info($"Exception occured: {ex.Message}");
                return new BadRequestObjectResult(ex);
            }
        }

        private static async Task ArchiveTeamAsync(string accessToken, ArchiveTeamRequest request)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            // Find groups with the specified SharePoint item ID
            var groupsToArchive = await graphClient.FindGroupsBySharePointItemIdAsync(request.SharePointItemId);

            foreach (var group in groupsToArchive.Value)
            {
                // Archive each matching team
                await graphClient.ArchiveTeamAsync(group.Id);
            }
        }
    }
}
