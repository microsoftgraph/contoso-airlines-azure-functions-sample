// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace CreateFlightTeam.Graph
{
    public class GraphService
    {
        private static readonly string graphEndpoint = "https://graph.microsoft.com/";

        private GraphServiceClient graphClient = null;

        private ILogger logger = null;

        public GraphService(ILogger log)
        {
            var clientCredential = new ClientCredential(Environment.GetEnvironmentVariable("AppSecret"));
            var authClient = ClientCredentialProvider.CreateClientApplication(
                Environment.GetEnvironmentVariable("AppId"), clientCredential, null,
                Environment.GetEnvironmentVariable("TenantId"));
            authClient.RedirectUri = Environment.GetEnvironmentVariable("RedirectUri");
            var authProvider = new ClientCredentialProvider(authClient);

            graphClient = new GraphServiceClient(authProvider);

            logger = log;
        }

        public async Task<List<string>> GetUserIds(List<string> pilots, List<string> flightAttendants, bool fullyQualified = false)
        {
            var userIds = new List<string>();

            // Look up each user to get their Id property
            foreach(var pilot in pilots)
            {
                var user = await GetUserByUpn(pilot);
                userIds.Add(fullyQualified ? $"{graphEndpoint}beta/users/{user.Id}" : user.Id);
            }

            foreach(var flightAttendant in flightAttendants)
            {
                var user = await GetUserByUpn(flightAttendant);
                userIds.Add(fullyQualified ? $"{graphEndpoint}beta/users/{user.Id}" : user.Id);
            }

            return userIds;
        }

        public async Task<User> GetUserByUpn(string upn)
        {
            try {
                var user = await graphClient.Users[upn].Request().GetAsync();
                return user;
            }
            catch (Exception ex) {
                logger.LogError($"Graph exception: {ex.Message}");
                return null;
            }
        }

        public async Task<User> GetUserByEmail(string email)
        {
            var results = await graphClient.Users.Request().Filter($"mail eq '{email}'").GetAsync();
            return results.CurrentPage[0];
        }

        public async Task<Group> CreateGroupAsync(Group group)
        {
            return await graphClient.Groups.Request().AddAsync(group);
        }

        public async Task CreateTeamAsync(string groupId, Team team)
        {
            var response = await graphClient.Groups[groupId].Team.Request().PutAsync(team);
        }

        public async Task<Invitation> CreateGuestInvitationAsync(Invitation invite)
        {
            return await graphClient.Invitations.Request().AddAsync(invite);
        }

        public async Task AddMemberAsync(string teamId, string userId, bool isOwner = false)
        {
            var user = new DirectoryObject { Id = userId };

            try
            {
                await graphClient.Groups[teamId].Members.References.Request().AddAsync(user);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Add member returned an error: {ex.Message}");
            }

            if (isOwner)
            {
                try
                {
                    await graphClient.Groups[teamId].Owners.References.Request().AddAsync(user);
                }
                catch (Exception ex)
                {
                    logger.LogWarning($"Add owner returned an error: {ex.Message}");
                }
            }
        }

        public async Task RemoveMemberAsync(string teamId, string userId, bool isOwner = false)
        {
            if (isOwner)
            {
                await graphClient.Groups[teamId].Owners[userId].Reference.Request().DeleteAsync();
            }

            await graphClient.Groups[teamId].Members[userId].Reference.Request().DeleteAsync();
        }

        public async Task<ITeamChannelsCollectionPage> GetTeamChannelsAsync(string teamId)
        {
            return await graphClient.Teams[teamId].Channels.Request().GetAsync();
        }

        public async Task<Channel> CreateTeamChannelAsync(string teamId, Channel channel)
        {
            return await graphClient.Teams[teamId].Channels.Request().AddAsync(channel);
        }

        public async Task AddAppToTeam(string teamId, TeamsAppInstallation app)
        {
            var response = await graphClient.Teams[teamId].InstalledApps.Request().AddAsync(app);
        }

        public async Task<Site> GetSharePointSiteAsync(string sitePath)
        {
            return await graphClient.Sites[sitePath].Request().GetAsync();
        }

        public async Task<Drive> GetSiteDriveAsync(string siteId, string driveName)
        {
            var drives = await graphClient.Sites[siteId].Drives.Request()
                .Top(50).GetAsync();

            foreach (var drive in drives.CurrentPage)
            {
                if (drive.Name == driveName)
                {
                    return drive;
                }
            }

            return null;
        }

        public async Task<PlannerPlan> CreatePlanAsync(PlannerPlan plan)
        {
            return await graphClient.Planner.Plans.Request().AddAsync(plan);
        }

        public async Task<PlannerBucket> CreateBucketAsync(PlannerBucket bucket)
        {
            return await graphClient.Planner.Buckets.Request().AddAsync(bucket);
        }

        public async Task<PlannerTask> CreatePlannerTaskAsync(PlannerTask task)
        {
            return await graphClient.Planner.Tasks.Request().AddAsync(task);
        }

        public async Task<Site> GetTeamSiteAsync(string teamId)
        {
            return await graphClient.Groups[teamId].Sites["root"].Request().GetAsync();
        }

        public async Task<List> CreateSharePointListAsync(string siteId, List list)
        {
            return await graphClient.Sites[siteId].Lists.Request().AddAsync(list);
        }

        public async Task ArchiveTeamAsync(string teamId)
        {
            await graphClient.Teams[teamId].Archive().Request().PostAsync();
        }

        public async Task AddTeamChannelTab(string teamId, string channelId, TeamsTab tab)
        {
            await graphClient.Teams[teamId].Channels[channelId].Tabs.Request().AddAsync(tab);
        }

        public async Task<ISiteListsCollectionPage> GetSiteListsAsync(string siteId)
        {
            return await graphClient.Sites[siteId].Lists.Request().GetAsync();
        }

        public async Task<SitePage> CreateSharePointPageAsync(string siteId, SitePage page)
        {
            return await graphClient.Sites[siteId].Pages.Request().AddAsync(page);
        }

        public async Task PublishSharePointPageAsync(string siteId, string pageId)
        {
            await graphClient.Sites[siteId].Pages[pageId].Publish().Request().PostAsync();
        }

        public async Task<Subscription> CreateListSubscription(string listUrl, string notificationUrl)
        {
            var newSubscription = new Subscription
            {
                ClientState = Guid.NewGuid().ToString(),
                Resource = listUrl,
                ChangeType = "updated",
                ExpirationDateTime = DateTime.UtcNow.AddDays(2),
                NotificationUrl = notificationUrl
            };

            return await graphClient.Subscriptions.Request().AddAsync(newSubscription);
        }

        public async Task RemoveAllSubscriptions()
        {
            var subscriptions = await graphClient.Subscriptions.Request().GetAsync();

            foreach (var subscription in subscriptions.CurrentPage)
            {
                await graphClient.Subscriptions[subscription.Id].Request().DeleteAsync();
            }
        }

        public async Task<IDriveItemDeltaCollectionPage> GetListDelta(string driveId, string deltaRequestUrl)
        {
            IDriveItemDeltaCollectionPage changes = null;
            if (string.IsNullOrEmpty(deltaRequestUrl))
            {
                if (string.IsNullOrEmpty(driveId))
                {
                    logger.LogError("GetListDelta: You must provide either a driveId or deltaRequestUrl");
                    return null;
                }

                // New delta request
                changes = await graphClient.Drives[driveId].Root.Delta().Request().GetAsync();
            }
            else
            {
                changes = new DriveItemDeltaCollectionPage();
                changes.InitializeNextPageRequest(graphClient, deltaRequestUrl);
                changes = await changes.NextPageRequest.GetAsync();
            }

            return changes;
        }

        public async Task<ListItem> GetDriveItemListItem(string driveId, string itemId)
        {
            try
            {
                return await graphClient.Drives[driveId].Items[itemId].ListItem.Request().GetAsync();
            }
            catch (Exception)
            {
                // When document is first created and no fields are filled in, this call
                // fails with a NotFound error
                return null;
            }
        }

        public async Task SendUserNotification(string userId, string title, string message)
        {
            // Check for a user token for this user ID
            // If we do not have one, it may be because the user has not
            // ever used the mobile app. In this case, do nothing
            var token = await AuthProvider.GetUserToken(userId);
            if (string.IsNullOrEmpty(token)) {
                return;
            }

            var notifGraphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async(requestMessage) => {
                        var userToken = await AuthProvider.GetUserToken(userId);
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", userToken);
                    }
                )
            );

            var notification = new Notification
            {
                TargetHostName = Environment.GetEnvironmentVariable("NotificationHostName"),
                ExpirationDateTime = DateTimeOffset.UtcNow.AddHours(2),
                DisplayTimeToLive = 30,
                Priority = Priority.High,
                GroupName = "FlightChanges",
                TargetPolicy = new TargetPolicyEndpoints
                {
                    PlatformTypes = new string[] { "ios" }
                },
                Payload = new PayloadTypes
                {
                    VisualContent = new VisualProperties
                    {
                        Title = title,
                        Body = message
                    }
                },
                AdditionalData = new Dictionary<string, object>()
            };

            notification.AdditionalData.Add("appNotificationId", Guid.NewGuid().ToString());

            try
            {
                await notifGraphClient.Me.Notifications.Request().AddAsync(notification);
            }
            catch (ServiceException ex)
            {
                logger.LogWarning($"Error sending notification to {userId}: {ex.Message}");
            }
        }

        #region Unused code
/*
        public async Task CreateChatMessageAsync(string teamId, string channelId, ChatMessage message)
        {
            var response = await MakeGraphCall(HttpMethod.Post, $"/teams/{teamId}/channels/{channelId}/messages", message);
        }

        public async Task<DriveItem> GetOneDriveItemAsync(string siteId, string itemPath)
        {
            var response = await MakeGraphCall(HttpMethod.Get, $"/sites/{siteId}/drive/{itemPath}");
            return JsonConvert.DeserializeObject<DriveItem>(await response.Content.ReadAsStringAsync());
        }

        public async Task<DriveItem> GetTeamOneDriveFolderAsync(string teamId, string folderName)
        {
            // Retry this call twice if it fails
            // There seems to be a delay between creating a Team and the drives being
            // fully created/enabled
            var response = await MakeGraphCall(HttpMethod.Get, $"/groups/{teamId}/drive/root:/{folderName}", retries: 3);
            return JsonConvert.DeserializeObject<DriveItem>(await response.Content.ReadAsStringAsync());
        }

        public async Task CopySharePointFileAsync(string siteId, string itemId, ItemReference target)
        {
            var copyPayload = new DriveItem
            {
                ParentReference = target
            };

            var response = await MakeGraphCall(HttpMethod.Post,
                $"/sites/{siteId}/drive/items/{itemId}/copy",
                copyPayload);
        }

        public async Task AddOpenExtensionToGroupAsync(string groupId, ProvisioningExtension extension)
        {
            var response = await MakeGraphCall(HttpMethod.Post, $"/groups/{groupId}/extensions", extension);
        }

        public async Task<SharePointList> GetSharePointListAsync(string siteId, string listName)
        {
            var response = await MakeGraphCall(HttpMethod.Get, $"/sites/{siteId}/lists?$top=50");
            var lists = JsonConvert.DeserializeObject<GraphCollection<SharePointList>>(await response.Content.ReadAsStringAsync());
            foreach(var list in lists.Value)
            {
                if (list.DisplayName == listName)
                {
                    return list;
                }
            }

            return null;
        }

        public async Task<GraphCollection<Group>> FindGroupsBySharePointItemIdAsync(int itemId)
        {
            var response = await MakeGraphCall(HttpMethod.Get, $"/groups?$filter={Group.SchemaExtensionName}/sharePointItemId  eq {itemId}");
            return JsonConvert.DeserializeObject<GraphCollection<Group>>(await response.Content.ReadAsStringAsync());
        }

        public async Task<GraphCollection<User>> GetGroupMembersAsync(string groupId)
        {
            var response = await MakeGraphCall(HttpMethod.Get, $"/groups/{groupId}/members");
            return JsonConvert.DeserializeObject<GraphCollection<User>>(await response.Content.ReadAsStringAsync());
        }

        public async Task SendNotification(Notification notification)
        {
            var response = await MakeGraphCall(HttpMethod.Post, "/me/notifications", notification);
        }

        public async Task<GraphCollection<Group>> GetAllGroupsAsync(string filter = null)
        {
            string query = string.IsNullOrEmpty(filter) ? string.Empty : $"?$filter={filter}";
            var response = await MakeGraphCall(HttpMethod.Get, $"/groups{query}");
            return JsonConvert.DeserializeObject<GraphCollection<Group>>(await response.Content.ReadAsStringAsync());
        }

        public async Task DeleteGroupAsync(string groupId)
        {
            var response = await MakeGraphCall(HttpMethod.Delete, $"/groups/{groupId}");
        }

        private async Task<HttpResponseMessage> MakeGraphCall(HttpMethod method, string uri, object body = null, int retries = 0, string version = "beta")
        {
            // Initialize retry delay to 3 secs
            int retryDelay = 3;

            string payload = string.Empty;

            if (body != null && (method != HttpMethod.Get || method != HttpMethod.Delete))
            {
                // Serialize the body
                payload = JsonConvert.SerializeObject(body, jsonSettings);
            }

            if (logger != null)
            {
                logger.LogInformation($"MakeGraphCall Request: {method} {uri}");
                logger.LogInformation($"MakeGraphCall Payload: {payload}");
            }

            do
            {
                var requestUrl = uri.StartsWith("https") ? uri : $"{graphEndpoint}{version}{uri}";
                // Create the request
                var request = new HttpRequestMessage(method, requestUrl);


                if (!string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                }

                // Send the request
                var response = await httpClient.SendAsync(request);

                if (!response.IsSuccessStatusCode)
                {
                    if (logger != null)
                        logger.LogInformation($"MakeGraphCall Error: {response.StatusCode}");
                    if (retries > 0)
                    {
                        if (logger != null)
                            logger.LogInformation($"MakeGraphCall Retrying after {retryDelay} seconds...({retries} retries remaining)");
                        Thread.Sleep(retryDelay * 1000);
                        // Double the retry delay for subsequent retries
                        retryDelay += retryDelay;
                    }
                    else
                    {
                        // No more retries, throw error
                        var error = await response.Content.ReadAsStringAsync();
                        throw new Exception(error);
                    }
                }
                else
                {
                    return response;
                }
            }
            while (retries-- > 0);

            return null;
        }
        */
        #endregion
    }
}
