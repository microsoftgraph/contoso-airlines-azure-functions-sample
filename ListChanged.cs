using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using CreateFlightTeam.Models;
using CreateFlightTeam.DocumentDB;
using CreateFlightTeam.Graph;
using System.Linq;
using CreateFlightTeam.Provisioning;

namespace CreateFlightTeam
{
    [StorageAccount("AzureWebJobsStorage")]
    public static class ListChanged
    {
        private static readonly string NotificationUrl =
            string.IsNullOrEmpty(Environment.GetEnvironmentVariable("NgrokProxy")) ?
                $"https://{Environment.GetEnvironmentVariable("WEBSITE_HOSTNAME")}/api/ListChanged" :
                $"{Environment.GetEnvironmentVariable("NgrokProxy")}/api/ListChanged";

        private static readonly string flightAdminSite = Environment.GetEnvironmentVariable("FlightAdminSite");
        private static readonly string flightList = Environment.GetEnvironmentVariable("FlightList");

        // This function implements a webhook
        // for a Graph subscription
        // https://docs.microsoft.com/graph/webhooks
        // This is called any time the Flights list is updated in
        // SharePoint
        [FunctionName("ListChanged")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            [Queue("syncqueue")] ICollector<ListChangedRequest> outputQueueMessage,
            ILogger log)
        {
            // Is this a validation request?
            if (req.Query.ContainsKey("validationToken"))
            {
                var validationToken = req.Query["validationToken"].ToString();
                log.LogInformation($"Validation request - Token : {validationToken}");
                return new OkObjectResult(validationToken);
            }

            // Get the notification payload and deserialize
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var request = JsonConvert.DeserializeObject<ListChangedRequest>(requestBody);

            // Add the notification to the queue
            outputQueueMessage.Add(request);

            // Return 202
            return new AcceptedResult();
        }

        // This function triggers on an item being added to the
        // queue by the ListChange function.
        // It does the processing of the notification
        [FunctionName("SyncList")]
        [Singleton]
        public static async Task SyncList(
            [QueueTrigger("syncqueue")] ListChangedRequest request,
            ILogger log)
        {
            log.LogInformation($"Received queue item: {JsonConvert.SerializeObject(request)}");

            DatabaseHelper.Initialize();
            AuthProvider.AzureLogger = log;

            // Validate the notification against the subscription
            var subscriptions = await DatabaseHelper.GetListSubscriptionsAsync(
                s => s.SubscriptionId == request.Changes[0].SubscriptionId);

            if (subscriptions.Count() > 1)
            {
                log.LogWarning($"There are {subscriptions.Count()} subscriptions in the database.");
            }

            var subscription = subscriptions.FirstOrDefault();

            if (subscription != null)
            {
                // Verify client state. If no match, no-op
                if (request.Changes[0].ClientState == subscription.ClientState)
                {
                    var graphClient = new GraphService(log);

                    // Process changes
                    var newDeltaLink = await ProcessDelta(graphClient, log, deltaLink: subscription.DeltaLink);

                    if (!string.IsNullOrEmpty(newDeltaLink))
                    {
                        subscription.DeltaLink = newDeltaLink;

                        // Update the subscription in the database with new delta link
                        await DatabaseHelper.UpdateListSubscriptionAsync(subscription.Id, subscription);
                    }
                }
            }
        }

        // This function is used to manually seed the flight team database
        // It will sync the database with the SharePoint list
        // and provision/update/remove any teams as needed
        [FunctionName("EnsureDatabase")]
        public static async Task EnsureDatabase(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            [Queue("syncqueue")] ICollector<ListChangedRequest> outputQueueMessage,
            ILogger log)
        {
            var graphClient = new GraphService(log);

            DatabaseHelper.Initialize();

            // Get the Flight Admin site
            var rootSite = await graphClient.GetSharePointSiteAsync("root");
            var adminSite = await graphClient.GetSharePointSiteAsync(
                $"{rootSite.SiteCollection.Hostname}:/sites/{flightAdminSite}");

            var drive = await graphClient.GetSiteDriveAsync(adminSite.Id, flightList);

            // Is there a subscription in the database?
            var subscriptions = await DatabaseHelper.GetListSubscriptionsAsync(s => s.Resource.Equals($"/drives/{drive.Id}/root"));
            var subscription = subscriptions.FirstOrDefault();

            if (subscription == null || subscription.IsExpired())
            {
                // Create a subscription
                var newSubscription = await graphClient.CreateListSubscription($"/drives/{drive.Id}/root", NotificationUrl);

                if (subscription == null)
                {
                    subscription = await DatabaseHelper.CreateListSubscriptionAsync(new ListSubscription
                    {
                        ClientState = newSubscription.ClientState,
                        Expiration = newSubscription.ExpirationDateTime.GetValueOrDefault().UtcDateTime,
                        Resource = $"/drives/{drive.Id}/root",
                        SubscriptionId = newSubscription.Id
                    });
                }
                else
                {
                    subscription.ClientState = newSubscription.ClientState;
                    subscription.Expiration = newSubscription.ExpirationDateTime.GetValueOrDefault().UtcDateTime;
                    subscription.SubscriptionId = newSubscription.Id;

                    subscription = await DatabaseHelper.UpdateListSubscriptionAsync(subscription.Id, subscription);
                }
            }

            string deltaLink = string.Empty;

            if (string.IsNullOrEmpty(subscription.DeltaLink))
            {
                deltaLink = await ProcessDelta(graphClient, log, driveId: drive.Id);
            }
            else
            {
                deltaLink = await ProcessDelta(graphClient, log, deltaLink: subscription.DeltaLink);
            }

            subscription.DeltaLink = deltaLink;
            await DatabaseHelper.UpdateListSubscriptionAsync(subscription.Id, subscription);
        }

        // This function is used to manually remove all subscriptions
        // and optionally clear the team database
        [FunctionName("Unsubscribe")]
        public static async Task Unsubscribe(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            DatabaseHelper.Initialize();

            var graphClient = new GraphService(log);

            await graphClient.RemoveAllSubscriptions();

            var subscriptions = await DatabaseHelper.GetListSubscriptionsAsync();
            foreach(var subscription in subscriptions)
            {
                await DatabaseHelper.DeleteListSubscriptionAsync(subscription.Id);
            }

            if (!string.IsNullOrEmpty(req.Query["deleteTeams"]))
            {
                var flightGroups = await graphClient.GetAllGroupsAsync("startswith(displayName, 'Flight ')");

                foreach(var group in flightGroups.CurrentPage)
                {
                    if (group.DisplayName.CompareTo(flightAdminSite) == 0 ||
                        group.DisplayName.CompareTo("Flight Admins") == 0 ||
                        group.DisplayName.CompareTo("Flight Attendants") == 0)
                    {
                        log.LogInformation($"Skipping required group: {group.DisplayName}");
                        continue;
                    }

                    try
                    {
                        await graphClient.DeleteTeamAsync(group.Id);
                    }
                    catch (Microsoft.Graph.ServiceException ex)
                    {
                        log.LogWarning($"Error deleting team ${group.Id}: {ex.Message}");
                    }

                    try
                    {
                        await graphClient.DeleteGroupAsync(group.Id);
                    }
                    catch (Microsoft.Graph.ServiceException ex)
                    {
                        log.LogWarning($"Error deleting group ${group.Id}: {ex.Message}");
                    }
                }

                var teams = await DatabaseHelper.GetFlightTeamsAsync();
                foreach (var team in teams)
                {
                    await DatabaseHelper.DeleteFlightTeamAsync(team.Id);
                }
            }
        }

        private static async Task<string> ProcessDelta(GraphService graphClient, ILogger log, string driveId = null, string deltaLink = null)
        {
            string deltaRequestUrl = deltaLink;

            TeamProvisioning.Initialize(graphClient, log);

            var delta = await graphClient.GetListDelta(driveId, deltaRequestUrl);

            foreach(var item in delta.CurrentPage)
            {
                await ProcessDriveItem(graphClient, item);
            }

            while(delta.NextPageRequest != null)
            {
                // There are more pages of results
                delta = await delta.NextPageRequest.GetAsync();

                foreach(var item in delta.CurrentPage)
                {
                    await ProcessDriveItem(graphClient, item);
                }
            }

            // Get the delta link
            object newDeltaLink;
            delta.AdditionalData.TryGetValue("@odata.deltaLink", out newDeltaLink);

            return newDeltaLink.ToString();
        }

        private static async Task ProcessDriveItem(GraphService graphClient, Microsoft.Graph.DriveItem item)
        {
            if (item.File != null)
            {
                // Query the database
                var teams = await DatabaseHelper.GetFlightTeamsAsync(f => f.SharePointListItemId.Equals(item.Id));
                var team = teams.FirstOrDefault();

                if (item.Deleted != null && team != null)
                {
                    // Remove the team
                    await TeamProvisioning.ArchiveTeamAsync(team);

                    // Remove the database item
                    await DatabaseHelper.DeleteFlightTeamAsync(team.Id);

                    return;
                }

                // Get the file's list data
                var listItem = await graphClient.GetDriveItemListItem(item.ParentReference.DriveId, item.Id);
                if (listItem == null) return;

                if (team == null)
                {
                    team = FlightTeam.FromListItem(item.Id, listItem);
                    if (team == null)
                    {
                        // Item was added to list but required metadata
                        // isn't filled in yet. No-op.
                        return;
                    }

                    // New item, provision team
                    team.TeamId = await TeamProvisioning.ProvisionTeamAsync(team);

                    await DatabaseHelper.CreateFlightTeamAsync(team);
                }
                else
                {
                    var updatedTeam = FlightTeam.FromListItem(item.Id, listItem);
                    updatedTeam.TeamId = team.TeamId;

                    // Existing item, process changes
                    await TeamProvisioning.UpdateTeamAsync(team, updatedTeam);
                    updatedTeam.Id = team.Id;
                    await DatabaseHelper.UpdateFlightTeamAsync(team.Id, updatedTeam);
                }
            }
        }
    }
}
