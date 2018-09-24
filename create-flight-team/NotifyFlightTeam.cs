// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
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
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class NotifyFlightTeam
    {
        private static readonly string notifAppId = Environment.GetEnvironmentVariable("NotificationAppId");
        private static readonly bool sendCrossDeviceNotifications = !string.IsNullOrEmpty(notifAppId);

        private static TraceWriter logger = null;

        [FunctionName("NotifyFlightTeam")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                log.Info($"Access token: {graphToken}");

                string requestBody = new StreamReader(req.Body).ReadToEnd();
                var request = JsonConvert.DeserializeObject<NotifyFlightTeamRequest>(requestBody);

                await NotifyTeamAsync(graphToken, request);

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

        private static async Task NotifyTeamAsync(string accessToken, NotifyFlightTeamRequest request)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            // Find groups with specified SharePoint item ID
            var groupsToNotify = await graphClient.FindGroupsBySharePointItemIdAsync(request.SharePointItemId);

            foreach (var group in groupsToNotify.Value)
            {
                // Post a Teams chat
                await PostTeamChatNotification(graphClient, group.Id, request.NewDepartureGate);

                if (sendCrossDeviceNotifications)
                {
                    // Get the group members
                    var members = await graphClient.GetGroupMembersAsync(group.Id);

                    // Send notification to each member
                    await SendNotificationAsync(graphClient, members.Value, group.DisplayName, request.NewDepartureGate);
                }
            }
        }

        private static async Task PostTeamChatNotification(GraphService graphClient, string groupId, string newDepartureGate)
        {
            // Get channels
            var channels = await graphClient.GetTeamChannelsAsync(groupId);

            // Create notification thread   
            var notificationThread = new ChatThread
            {
                RootMessage = new ChatMessage
                {
                    Body = new ItemBody { Content = $"Your flight will now depart from gate {newDepartureGate}" }
                }
            };

            // Post to all channels
            foreach (var channel in channels.Value)
            {
                await graphClient.CreateChatThreadAsync(groupId, channel.Id, notificationThread);
            }
        }

        private static async Task SendNotificationAsync(GraphService graphClient, List<User> users, string groupName, string newDepartureGate)
        {
            // Ideally loop through all the members here and send each a notification
            // The notification API is currently limited to only send to the logged-in user
            // So to do this, would need to manage tokens for each user.
            // For now, just send to the authenticated user.
            var notification = new Notification
            {
                TargetHostName = notifAppId,
                AppNotificationId = "testDirectToastNotification",
                GroupName = "TestGroup",
                ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(1).ToUniversalTime(),
                Priority = "High",
                DisplayTimeToLive = 30,
                Payload = new NotificationPayload
                {
                    VisualContent = new NotificationVisualContent
                    {
                        Title = $"{groupName} gate change",
                        Body = $"Departure gate has been changed to {newDepartureGate}"
                    }
                },
                TargetPolicy = new NotificationTargetPolicy
                {
                    PlatformTypes = new string[] { "windows", "android", "ios" }
                }
            };

            await graphClient.SendNotification(notification);
        }
    }
}
