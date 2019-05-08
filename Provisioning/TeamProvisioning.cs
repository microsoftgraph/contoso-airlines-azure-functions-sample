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
        private static readonly string webPartId = Environment.GetEnvironmentVariable("WebPartId");

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

            await AddFlightToCalendars(flightTeam);

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

            bool isCrewChanged = false;

            // Add new pilots
            var newPilots = updatedTeam.Pilots.Except(originalTeam.Pilots);
            foreach (var pilot in newPilots)
            {
                isCrewChanged = true;
                var pilotUser = await graphClient.GetUserByUpn(pilot);
                await graphClient.AddMemberAsync(originalTeam.TeamId, pilotUser.Id);
            }

            if (newPilots.Count() > 0)
            {
                await TeamProvisioning.AddFlightToCalendars(updatedTeam, newPilots.ToList());
            }

            // Remove any removed pilots
            var removedPilots = originalTeam.Pilots.Except(updatedTeam.Pilots);
            foreach (var pilot in removedPilots)
            {
                isCrewChanged = true;
                var pilotUser = await graphClient.GetUserByUpn(pilot);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, pilotUser.Id);
            }

            if (removedPilots.Count() > 0)
            {
                await TeamProvisioning.RemoveFlightFromCalendars(removedPilots.ToList(), updatedTeam.FlightNumber);
            }

            // Add new flight attendants
            var newFlightAttendants = updatedTeam.FlightAttendants.Except(originalTeam.FlightAttendants);
            foreach (var attendant in newFlightAttendants)
            {
                isCrewChanged = true;
                var attendantUser = await graphClient.GetUserByUpn(attendant);
                await graphClient.AddMemberAsync(originalTeam.TeamId, attendantUser.Id);
            }

            if (newFlightAttendants.Count() > 0)
            {
                await TeamProvisioning.AddFlightToCalendars(updatedTeam, newFlightAttendants.ToList());
            }

            // Remove any removed flight attendants
            var removedFlightAttendants = originalTeam.FlightAttendants.Except(updatedTeam.FlightAttendants);
            foreach (var attendant in removedFlightAttendants)
            {
                isCrewChanged = true;
                var attendantUser = await graphClient.GetUserByUpn(attendant);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, attendantUser.Id);
            }

            if (removedFlightAttendants.Count() > 0)
            {
                await TeamProvisioning.RemoveFlightFromCalendars(removedFlightAttendants.ToList(), updatedTeam.FlightNumber);
            }

            // Swap out catering liaison if needed
            if (updatedTeam.CateringLiaison != null &&
                !updatedTeam.CateringLiaison.Equals(originalTeam.CateringLiaison))
            {
                var oldCateringLiaison = await graphClient.GetUserByEmail(originalTeam.CateringLiaison);
                await graphClient.RemoveMemberAsync(originalTeam.TeamId, oldCateringLiaison.Id);
                await AddGuestUser(originalTeam.TeamId, updatedTeam.CateringLiaison);
            }

            // Check for changes to gate, time
            bool isGateChanged = updatedTeam.DepartureGate != originalTeam.DepartureGate;
            bool isDepartureTimeChanged = updatedTeam.DepartureTime != originalTeam.DepartureTime;

            List<string> crew = null;
            string newGate = null;

            if (isCrewChanged || isGateChanged || isDepartureTimeChanged)
            {
                crew = await graphClient.GetUserIds(updatedTeam.Pilots, updatedTeam.FlightAttendants);
                newGate = isGateChanged ? updatedTeam.DepartureGate : null;

                logger.LogInformation("Updating flight in crew members' calendars");

                if (isDepartureTimeChanged)
                {
                    await TeamProvisioning.UpdateFlightInCalendars(crew, updatedTeam.FlightNumber, updatedTeam.DepartureGate, updatedTeam.DepartureTime);
                }
                else
                {
                    await TeamProvisioning.UpdateFlightInCalendars(crew, updatedTeam.FlightNumber, updatedTeam.DepartureGate);
                }
            }

            if (isGateChanged || isDepartureTimeChanged)
            {
                var localTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                string newDepartureTime = isDepartureTimeChanged ? TimeZoneInfo.ConvertTime(updatedTeam.DepartureTime, localTimeZone).ToString("g") : null;

                logger.LogInformation("Sending notification to crew members' devices");
                await TeamProvisioning.SendDeviceNotifications(crew, updatedTeam.FlightNumber,
                    newGate, newDepartureTime);
            }
        }

        public static async Task ArchiveTeamAsync(FlightTeam team)
        {
            var crew = await graphClient.GetUserIds(team.Pilots, team.FlightAttendants);

            // Remove event from crew calendars
            await TeamProvisioning.RemoveFlightFromCalendars(crew, team.FlightNumber);
            // Archive team
            try
            {
                await graphClient.ArchiveTeamAsync(team.TeamId);
            }
            catch (ServiceException ex)
            {
                logger.LogInformation($"Attempt to archive team failed: {ex.Message}");
            }
        }

        private static async Task<Group> CreateUnifiedGroupAsync(FlightTeam flightTeam)
        {
            // Initialize members list with pilots and flight attendants
            var members = await graphClient.GetUserIds(flightTeam.Pilots, flightTeam.FlightAttendants, true);

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

        private static async Task<List> CreateChallengingPassengersListAsync(string groupId, string channelId)
        {
            int retries = 3;

            while (retries > 0)
            {
                try
                {
                    // Get the team site
                    var teamSite = await graphClient.GetTeamSiteAsync(groupId);

                    var challengingPassengers = new List
                    {
                        DisplayName = "Challenging Passengers",
                        Columns = new ListColumnsCollectionPage {
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

                    logger.LogInformation("Created challenging passenger list");
                    return createdList;
                }
                catch (ServiceException ex)
                {
                    logger.LogWarning($"CreateChallengingPassengersListAsync error: {ex.Message}");
                    retries--;
                    logger.LogWarning($"{retries} retries remaining");
                }
            }

            return null;
        }

        private static async Task CreateSharePointPageAsync(string groupId, string channelId, float flightNumber)
        {
            try
            {
            // Get the team site
                var teamSite = await graphClient.GetTeamSiteAsync(groupId);
                logger.LogInformation("Got team site");

                // Initialize page
                var sharePointPage = new SitePage
                {
                    Name = "Crew.aspx",
                    Title = $"Flight {flightNumber} Crew"
                };

                var webParts = new List<WebPart>
                {
                    new WebPart
                    {
                        Type = webPartId,
                        Data = new SitePageData
                        {
                            AdditionalData = new Dictionary<string, object>
                            {
                                { "dataVersion", "1.0"},
                                { "properties", new Dictionary<string, object>
                                    {
                                        { "description", "CrewBadges" }
                                    }
                                }
                            }
                        }
                    }
                };

                sharePointPage.WebParts = webParts;

                var createdPage = await graphClient.CreateSharePointPageAsync(teamSite.Id, sharePointPage);
                logger.LogInformation("Created crew page");

                // Publish the page
                await graphClient.PublishSharePointPageAsync(teamSite.Id, createdPage.Id);
                var pageUrl = createdPage.WebUrl.StartsWith("https") ? createdPage.WebUrl :
                    $"{teamSite.WebUrl}/{createdPage.WebUrl}";

                logger.LogInformation("Published crew page");

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

                logger.LogInformation("Added crew page as Teams tab");
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Failed to create crew page: ${ex.ToString()}");
            }
        }

        private static async Task AddFlightToCalendars(FlightTeam flightTeam, List<string> usersToAdd = null)
        {
            // Get all flight members
            var allCrewIds = await graphClient.GetUserIds(flightTeam.Pilots, flightTeam.FlightAttendants);

            // Initialize flight event
            var flightEvent = new Event
            {
                Subject = $"Flight {flightTeam.FlightNumber}",
                Location = new Location
                {
                    DisplayName = flightTeam.Description
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = flightTeam.DepartureTime.ToString("s"),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = flightTeam.DepartureTime.AddHours(4).ToString("s"),
                    TimeZone = "UTC"
                },
                Categories = new string[] { "Assigned Flight" },
                Extensions = new EventExtensionsCollectionPage()
            };

            var flightExtension = new OpenTypeExtension
            {
                ODataType = "microsoft.graph.openTypeExtension",
                ExtensionName = "com.contoso.flightData",
                AdditionalData = new Dictionary<string, object>()
            };

            flightExtension.AdditionalData.Add("departureGate", flightTeam.DepartureGate);
            flightExtension.AdditionalData.Add("crewMembers", allCrewIds);

            flightEvent.Extensions.Add(flightExtension);

            if (usersToAdd == null)
            {
                usersToAdd = allCrewIds;
            }

            foreach (var userId in usersToAdd)
            {
                //var user = await graphClient.GetUserByUpn(userId);

                await graphClient.CreateEventInUserCalendar(userId, flightEvent);
            }
        }

        private static async Task UpdateFlightInCalendars(List<string> crewMembers, int flightNumber, string departureGate, DateTime? newDepartureTime = null)
        {
            foreach (var userId in crewMembers)
            {
                // Get the event
                var matchingEvents = await graphClient.GetEventsInUserCalendar(userId,
                    $"categories/any(a:a eq 'Assigned Flight') and subject eq 'Flight {flightNumber}'");

                var flightEvent = matchingEvents.CurrentPage.FirstOrDefault();
                if (flightEvent != null)
                {
                    if (newDepartureTime != null)
                    {
                        flightEvent.Start.DateTime = newDepartureTime?.ToString("s");
                        flightEvent.End.DateTime = newDepartureTime?.AddHours(4).ToString("s");
                        await graphClient.UpdateEventInUserCalendar(userId, flightEvent);
                    }

                    var flightExtension = new OpenTypeExtension
                    {
                        ODataType = "microsoft.graph.openTypeExtension",
                        ExtensionName = "com.contoso.flightData",
                        AdditionalData = new Dictionary<string, object>()
                    };

                    flightExtension.AdditionalData.Add("departureGate", departureGate);
                    flightExtension.AdditionalData.Add("crewMembers", crewMembers);

                    await graphClient.UpdateFlightExtension(userId, flightEvent.Id, flightExtension);
                }
            }
        }

        private static async Task RemoveFlightFromCalendars(List<string> usersToRemove, int flightNumber)
        {
            foreach (var userId in usersToRemove)
            {
                logger.LogInformation($"Deleting flight from ${userId}");
                // Get the event
                try
                {
                    var matchingEvents = await graphClient.GetEventsInUserCalendar(userId,
                        $"categories/any(a:a eq 'Assigned Flight') and subject eq 'Flight {flightNumber}'");

                    var flightEvent = matchingEvents.CurrentPage.FirstOrDefault();
                    if (flightEvent != null)
                    {
                        await graphClient.DeleteEventInUserCalendar(userId, flightEvent.Id);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning($"Delete event returned an error: {ex.Message}");
                }
            }
        }

        private static async Task SendDeviceNotifications(List<string> crewMembers, int flightNumber, string newGate = null, string newDepartureTime = null)
        {
            string notificationText = string.Empty;

            if (!string.IsNullOrEmpty(newGate))
            {
                notificationText = $"New Departure Gate: {newGate}";
            }

            if (!string.IsNullOrEmpty(newDepartureTime))
            {
                notificationText = $"{(string.IsNullOrEmpty(notificationText) ? "" : notificationText + "\n")}New Departure Time: {newDepartureTime}";
            }

            if (!string.IsNullOrEmpty(notificationText))
            {
                foreach(var userId in crewMembers)
                {
                    await graphClient.SendUserNotification(userId, $"Flight {flightNumber} Update", notificationText);
                }
            }
        }

        private static string GetTimestamp()
        {
            var now = DateTime.Now;
            return $"{now.Hour}{now.Minute}{now.Second}";
        }
    }
}
