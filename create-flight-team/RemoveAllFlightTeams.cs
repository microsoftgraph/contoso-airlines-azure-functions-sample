
using create_flight_team.Graph;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class RemoveAllFlightTeams
    {
        private static TraceWriter logger = null;

        [FunctionName("RemoveAllFlightTeams")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                log.Info($"Access token: {graphToken}");

                await RemoveAllTeamsAsync(graphToken);

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

        public static async Task RemoveAllTeamsAsync(string accessToken)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            bool more = true;

            do
            {
                var groups = await graphClient.GetAllGroupsAsync("startswith(displayName, 'Flight')");

                more = groups.Value.Count > 0;

                foreach (var group in groups.Value)
                {
                    if (group.DisplayName != "Flight Admin")
                    {
                        logger.Info($"Deleting team {group.DisplayName}");

                        // Archive the team
                        try
                        {
                            await graphClient.ArchiveTeamAsync(group.Id);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("ItemNotFound"))
                            {
                                logger.Info("No team found");
                            }
                            else { throw ex; }
                        }

                        // Delete the group
                        await graphClient.DeleteGroupAsync(group.Id);
                    }
                }
            }
            while (more);
        }
    }
}
