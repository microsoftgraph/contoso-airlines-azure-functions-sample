using CreateFlightTeam.Graph;
using CreateFlightTeam.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CreateFlightTeam.Provisioning
{
    public static class TeamProvisioning
    {
        private static readonly string teamAppId = Environment.GetEnvironmentVariable("TeamAppToInstall");
        private static readonly string tenantName = Environment.GetEnvironmentVariable("TenantName");

        private static GraphService graphClient;
        private static ILogger logger;

        public static void Initialize(GraphService client, ILogger log)
        {
            graphClient = client;
            logger = log;
        }

        public static async Task<string> ProvisionTeamAsync(FlightTeam flightTeam)
        {
            // Create the unified group
            var group = await CreateUnifiedGroupAsync(flightTeam);

            // Create the team in the group
            var teamChannel = await InitializeTeamInGroupAsync(group.Id,
                $"Welcome to Flight {flightTeam.FlightNumber}!");

            // Create Planner plan and tasks
            // TODO: Disabled because you cannot create planner plans with app-only token
            // await CreatePreflightPlanAsync(group.Id, teamChannel.Id, flightTeam.DepartureTime);

            // Create SharePoint list
            await CreateChallengingPassengersListAsync(group.Id, teamChannel.Id);

            // Create SharePoint page
            await CreateSharePointPageAsync(group.Id, teamChannel.Id, flightTeam.FlightNumber);

            return group.Id;
        }

        public static async Task UpdateTeamAsync(FlightTeam originalTeam, FlightTeam updatedTeam)
        {
            // Look for changes that require an update via Graph
            // Did the admin change?
            var admin = await graphClient.GetUserByUpn(updatedTeam.Admin);
            updatedTeam.Admin = admin.Id;
            if (!admin.Id.Equals(originalTeam.Admin))
            {
                // Add new owner
                await graphClient.AddMemberAsync(originalTeam.TeamId, admin.Id, true);
                // Remove old owner
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, admin.Id, true);
            }

            // Add new pilots
            var newPilots = updatedTeam.Pilots.Except(originalTeam.Pilots);
            foreach (var pilot in newPilots)
            {
                var pilotUser = await graphClient.GetUserByUpn(pilot);
                await graphClient.AddMemberAsync(originalTeam.TeamId, pilotUser.Id);
            }

            // Remove any removed pilots
            var removedPilots = originalTeam.Pilots.Except(updatedTeam.Pilots);
            foreach (var pilot in removedPilots)
            {
                var pilotUser = await graphClient.GetUserByUpn(pilot);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, pilotUser.Id);
            }

            // Add new flight attendants
            var newFlightAttendants = updatedTeam.FlightAttendants.Except(originalTeam.FlightAttendants);
            foreach (var attendant in newFlightAttendants)
            {
                var attendantUser = await graphClient.GetUserByUpn(attendant);
                await graphClient.AddMemberAsync(originalTeam.TeamId, attendantUser.Id);
            }

            // Remove any removed flight attendants
            var removedFlightAttendants = originalTeam.FlightAttendants.Except(updatedTeam.FlightAttendants);
            foreach (var attendant in removedFlightAttendants)
            {
                var attendantUser = await graphClient.GetUserByUpn(attendant);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, attendantUser.Id);
            }

            // Swap out catering liaison if needed
            if (!updatedTeam.CateringLiaison.Equals(originalTeam.CateringLiaison))
            {
                var oldCateringLiaison = await graphClient.GetUserByEmail(originalTeam.CateringLiaison);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, oldCateringLiaison.Id);
                await AddGuestUser(originalTeam.TeamId, updatedTeam.CateringLiaison);
            }
        }

        public static async Task ArchiveTeamAsync(string teamId)
        {
            // Archive each matching team
            await graphClient.ArchiveTeamAsync(teamId);
        }

        private static async Task<Group> CreateUnifiedGroupAsync(FlightTeam flightTeam)
        {
            // Initialize members list with pilots and flight attendants
            var members = await graphClient.GetUserIds(flightTeam.Pilots, flightTeam.FlightAttendants);

            // Add admin and me as members
            members.Add($"https://graph.microsoft.com/beta/users/{flightTeam.Admin}");

            // Create owner list
            var owners = new List<string>() { $"https://graph.microsoft.com/beta/users/{flightTeam.Admin}" };

            // Create the group
            var flightGroup = new Group
            {
                DisplayName = $"Flight {flightTeam.FlightNumber}",
                Description = flightTeam.Description,
                Visibility = "Private",
                MailEnabled = true,
                MailNickname = $"flight{flightTeam.FlightNumber}{GetTimestamp()}",
                GroupTypes = new string[] { "Unified" },
                SecurityEnabled = false,
                AdditionalData = new Dictionary<string, object>()
            };

            flightGroup.AdditionalData.Add("members@odata.bind", members);
            flightGroup.AdditionalData.Add("owners@odata.bind", owners);

            var createdGroup = await graphClient.CreateGroupAsync(flightGroup);
            logger.LogInformation("Created group");

            if (!string.IsNullOrEmpty(flightTeam.CateringLiaison))
            {
                await AddGuestUser(createdGroup.Id, flightTeam.CateringLiaison);
            }

            return createdGroup;
        }

        private static async Task AddGuestUser(string groupId, string email)
        {
            // Add catering liaison as a guest
            var guestInvite = new Invitation
            {
                InvitedUserEmailAddress = email,
                InviteRedirectUrl = "https://teams.microsoft.com",
                SendInvitationMessage = true
            };

            var createdInvite = await graphClient.CreateGuestInvitationAsync(guestInvite);

            // Add guest user to team
            await graphClient.AddMemberAsync(groupId, createdInvite.InvitedUser.Id);
            logger.LogInformation("Added guest user");
        }

        private static async Task<Channel> InitializeTeamInGroupAsync(string groupId, string welcomeMessage)
        {
            // Create the team
            var team = new Team
            {
                GuestSettings = new TeamGuestSettings
                {
                    AllowCreateUpdateChannels = false,
                    AllowDeleteChannels = false
                }
            };

            await graphClient.CreateTeamAsync(groupId, team);
            logger.LogInformation("Created team");

            // Get channels
            var channels = await graphClient.GetTeamChannelsAsync(groupId);

            // Get "General" channel. Since it is created by default and is the only
            // channel after creation, just get the first result.
            var generalChannel = channels.CurrentPage.First();

            //// Create welcome message (new thread)
            //var welcomeThread = new ChatThread
            //{
            //    RootMessage = new ChatMessage
            //    {
            //        Body = new ItemBody { Content = welcomeMessage }
            //    }
            //};

            //await graphClient.CreateChatThreadAsync(groupId, generalChannel.Id, welcomeThread);
            //logger.LogInformation("Posted welcome message");

            // Provision pilot channel
            var pilotChannel = new Channel
            {
                DisplayName = "Pilots",
                Description = "Discussion about flightpath, weather, etc."
            };

            await graphClient.CreateTeamChannelAsync(groupId, pilotChannel);
            logger.LogInformation("Created Pilots channel");

            // Provision flight attendants channel
            var flightAttendantsChannel = new Channel
            {
                DisplayName = "Flight Attendants",
                Description = "Discussion about duty assignments, etc."
            };

            await graphClient.CreateTeamChannelAsync(groupId, flightAttendantsChannel);
            logger.LogInformation("Created FA channel");

            // Add the requested team app
            if (!string.IsNullOrEmpty(teamAppId))
            {
                var teamsApp = new TeamsAppInstallation
                {
                    AdditionalData = new Dictionary<string, object>()
                };

                teamsApp.AdditionalData.Add("teamsApp@odata.bind",
                    $"https://graph.microsoft.com/beta/appCatalogs/teamsApps/{teamAppId}");

                await graphClient.AddAppToTeam(groupId, teamsApp);
            }
            logger.LogInformation("Added app to team");

            // Return the general channel
            return generalChannel;
        }

        private static async Task CreatePreflightPlanAsync(string groupId, string channelId, DateTimeOffset departureTime)
        {
            // Create a "Pre-flight checklist" plan
            var preFlightCheckList = new PlannerPlan
            {
                Title = "Pre-flight Checklist",
                Owner = groupId
            };

            var createdPlan = await graphClient.CreatePlanAsync(preFlightCheckList);
            logger.LogInformation("Create plan");

            // Create buckets
            var toDoBucket = new PlannerBucket
            {
                Name = "To Do",
                PlanId = createdPlan.Id
            };

            var createdToDoBucket = await graphClient.CreateBucketAsync(toDoBucket);

            var completedBucket = new PlannerBucket
            {
                Name = "Completed",
                PlanId = createdPlan.Id
            };

            var createdCompletedBucket = await graphClient.CreateBucketAsync(completedBucket);

            // Create tasks in to-do bucket
            var preFlightInspection = new PlannerTask
            {
                Title = "Perform pre-flight inspection of aircraft",
                PlanId = createdPlan.Id,
                BucketId = createdToDoBucket.Id,
                DueDateTime = departureTime.ToUniversalTime()
            };

            await graphClient.CreatePlannerTaskAsync(preFlightInspection);

            var ensureFoodBevStock = new PlannerTask
            {
                Title = "Ensure food and beverages are fully stocked",
                PlanId = createdPlan.Id,
                BucketId = createdToDoBucket.Id,
                DueDateTime = departureTime.ToUniversalTime()
            };

            await graphClient.CreatePlannerTaskAsync(ensureFoodBevStock);

            // Add planner tab to General channel
            var plannerTab = new TeamsTab
            {
                Name = "Pre-flight Checklist",
                TeamsAppId = "com.microsoft.teamspace.tab.planner",
                Configuration = new TeamsTabConfiguration
                {
                    EntityId = createdPlan.Id,
                    ContentUrl = $"https://tasks.office.com/{tenantName}/Home/PlannerFrame?page=7&planId={createdPlan.Id}&auth_pvr=Orgid&auth_upn={{upn}}&mkt={{locale}}",
                    RemoveUrl = $"https://tasks.office.com/{tenantName}/Home/PlannerFrame?page=13&planId={createdPlan.Id}&auth_pvr=Orgid&auth_upn={{upn}}&mkt={{locale}}",
                    WebsiteUrl = $"https://tasks.office.com/{tenantName}/Home/PlanViews/{createdPlan.Id}"
                }
            };

            await graphClient.AddTeamChannelTab(groupId, channelId, plannerTab);
        }

        private static async Task<List> CreateChallengingPassengersListAsync(string groupId, string channelId)
        {
            // Get the team site
            var teamSite = await graphClient.GetTeamSiteAsync(groupId);

            var challengingPassengers = new List
            {
                DisplayName = "Challenging Passengers",
                /*
                Columns = new List<ColumnDefinition>()
                {
                    new ColumnDefinition
                    {
                        Name = "Name",
                        Text = new TextColumn()
                    },
                    new ColumnDefinition
                    {
                        Name = "SeatNumber",
                        Text = new TextColumn()
                    },
                    new ColumnDefinition
                    {
                        Name = "Notes",
                        Text = new TextColumn()
                    }
                }
                 */
            };

            // Create the list
            var createdList = await graphClient.CreateSharePointListAsync(teamSite.Id, challengingPassengers);

            // Add the list as a team tab
            /*
            var listTab = new TeamsChannelTab
            {
                Name = "Challenging Passengers",
                TeamsAppId = "com.microsoft.teamspace.tab.web",
                Configuration = new TeamsChannelTabConfiguration
                {
                    ContentUrl = createdList.WebUrl,
                    WebsiteUrl = createdList.WebUrl
                }
            };

            await graphClient.AddTeamChannelTab(groupId, channelId, listTab);
             */

            return createdList;
        }

        private static async Task CreateSharePointPageAsync(string groupId, string channelId, float flightNumber)
        {
            // Get the team site
            var teamSite = await graphClient.GetTeamSiteAsync(groupId);

            // Get the site lists
            var siteLists = await graphClient.GetSiteListsAsync(teamSite.Id);

            // Initialize page
            var sharePointPage = new SitePage
            {
                Name = "TeamPage.aspx",
                Title = $"Flight {flightNumber}"
            };

            var webParts = new List<WebPart>();

            foreach (var list in siteLists.CurrentPage)
            {
                bool isDocLibrary = list.DisplayName == "Documents";

                var webPart = new Microsoft.Graph.WebPart
                {
                    Type = SharePointWebPart.ListWebPart,
                    Data = new SitePageData
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "dataVersion", "1.0"},
                            { "properties", new Dictionary<string, object>
                                {
                                    { "isDocumentLibrary", isDocLibrary },
                                    { "selectedListId", list.Id }
                                }
                            }
                        }
                    }
                };

                webParts.Add(webPart);
            }

            sharePointPage.WebParts = webParts;

            var createdPage = await graphClient.CreateSharePointPageAsync(teamSite.Id, sharePointPage);

            // Publish the page
            await graphClient.PublishSharePointPageAsync(teamSite.Id, createdPage.Id);
            var pageUrl = createdPage.WebUrl.StartsWith("https") ? createdPage.WebUrl :
                $"{teamSite.WebUrl}/{createdPage.WebUrl}";

            // Add the list as a team tab
            var pageTab = new TeamsTab
            {
                Name = createdPage.Title,
                TeamsAppId = "com.microsoft.teamspace.tab.web",
                Configuration = new TeamsTabConfiguration
                {
                    ContentUrl = pageUrl,
                    WebsiteUrl = pageUrl
                }
            };

            await graphClient.AddTeamChannelTab(groupId, channelId, pageTab);
        }

        private static string GetTimestamp()
        {
            var now = DateTime.Now;
            return $"{now.Hour}{now.Minute}{now.Second}";
        }
    }
}
