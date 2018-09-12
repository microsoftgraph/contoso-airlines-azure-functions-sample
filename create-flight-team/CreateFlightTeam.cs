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
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class CreateFlightTeam
    {
        private static readonly string teamAppId = Environment.GetEnvironmentVariable("TeamAppToInstall");
        private static readonly string flightAdminSite = Environment.GetEnvironmentVariable("FlightAdminSite");
        private static readonly string flightLogFile = Environment.GetEnvironmentVariable("FlightLogFile");

        private static TraceWriter logger = null;

        [FunctionName("CreateFlightTeam")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                log.Info($"Access token: {graphToken}");

                string requestBody = new StreamReader(req.Body).ReadToEnd();
                var request = JsonConvert.DeserializeObject<CreateFlightTeamRequest>(requestBody);

                await ProvisionTeam(graphToken, request);

                return new OkObjectResult(new CreateFlightTeamResponse { Result = "success" });
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
                return new BadRequestObjectResult(new CreateFlightTeamResponse { Result = "failed", Details = ex });
            }
        }

        

        private static async Task ProvisionTeam(string accessToken, CreateFlightTeamRequest request)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            // Create the unified group
            var group = await CreateUnifiedGroupAsync(graphClient, request);

            // Create the team in the group
            await InitializeTeamInGroupAsync(graphClient, group.Id, 
                $"Welcome to Flight {request.FlightNumber}!");

            // Copy flight log template to team files
            await CopyFlightLogToTeamFilesAsync(graphClient, group.Id);

            // Create Planner plan and tasks
            await CreatePreflightPlanAsync(graphClient, group.Id, request.DepartureTime);

            // Create SharePoint list
            await CreateChallengingPassengersListAsync(graphClient, group.Id);

            // Create SharePoint page
            // TODO
        }

        private static async Task<Group> CreateUnifiedGroupAsync(GraphService graphClient, CreateFlightTeamRequest request)
        {
            // Initialize members list with pilots and flight attendants
            var members = await graphClient.GetUserIds(request.Pilots, request.FlightAttendants);

            // Handle admins
            var admin = await graphClient.GetUserByUpn(request.Admin);
            var me = await graphClient.GetMe();

            // Add admin and me as members
            members.Add($"https://graph.microsoft.com/beta/users/{admin.Id}");
            members.Add($"https://graph.microsoft.com/beta/users/{me.Id}");

            // Create owner list
            var owners = new List<string>()
            {
                $"https://graph.microsoft.com/beta/users/{admin.Id}",
                $"https://graph.microsoft.com/beta/users/{me.Id}"
            };

            // Create the group
            var flightGroup = new Group
            {
                DisplayName = $"Flight {request.FlightNumber}",
                Description = request.Description,
                Visibility = "Private",
                MailEnabled = true,
                MailNickname = $"flight{request.FlightNumber}{GetTimestamp()}",
                GroupTypes = new string[] { "Unified" },
                SecurityEnabled = false,
                Members = members.Distinct().ToList(),
                Owners = owners.Distinct().ToList()
            };

            var createdGroup = await graphClient.CreateGroupAsync(flightGroup);
            logger.Info("Created group");

            // Add catering liaison as a guest
            var guestInvite = new Invitation
            {
                InvitedUserEmailAddress = request.CateringLiaison,
                InviteRedirectUrl = "https://teams.microsoft.com",
                SendInvitationMessage = true
            };

            var createdInvite = await graphClient.CreateGuestInvitationAsync(guestInvite);

            // Add guest user to team
            await graphClient.AddMemberAsync(createdGroup.Id, createdInvite.InvitedUser.Id);
            logger.Info("Added guest user");

            return createdGroup;
        }

        private static async Task InitializeTeamInGroupAsync(GraphService graphClient, string groupId, string welcomeMessage)
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
            logger.Info("Created team");

            // Get channels
            var channels = await graphClient.GetTeamChannelsAsync(groupId);

            // Get "General" channel. Since it is created by default and is the only
            // channel after creation, just get the first result.
            var generalChannel = channels.Value.First();

            // Create welcome message (new thread)
            var welcomeThread = new ChatThread
            {
                RootMessage = new ChatMessage
                {
                    Body = new ItemBody { Content = welcomeMessage }
                }
            };

            await graphClient.CreateChatThreadAsync(groupId, generalChannel.Id, welcomeThread);
            logger.Info("Posted welcome message");

            // Provision pilot channel
            var pilotChannel = new Channel
            {
                DisplayName = "Pilots",
                Description = "Discussion about flightpath, weather, etc."
            };

            await graphClient.CreateTeamChannelAsync(groupId, pilotChannel);
            logger.Info("Created Pilots channel");

            // Provision flight attendants channel
            var flightAttendantsChannel = new Channel
            {
                DisplayName = "Flight Attendants",
                Description = "Discussion about duty assignments, etc."
            };

            await graphClient.CreateTeamChannelAsync(groupId, flightAttendantsChannel);
            logger.Info("Created FA channel");

            // Add the requested team app
            if (!string.IsNullOrEmpty(teamAppId))
            {
                var teamsApp = new TeamsApp
                {
                    Id = teamAppId
                };

                await graphClient.AddAppToTeam(groupId, teamsApp);
            }
            logger.Info("Added app to team");
        }

        private static async Task CopyFlightLogToTeamFilesAsync(GraphService graphClient, string groupId)
        {
            // Upload flight log to team files
            // Get root site to determine SP host name
            var rootSite = await graphClient.GetSharePointSiteAsync("root");

            // Get flight admin site
            var adminSite = await graphClient.GetSharePointSiteAsync(
                $"{rootSite.SiteCollection.Hostname}:/sites/{flightAdminSite}");
            logger.Info("Got flight admin site");

            // Get the flight log document
            var flightLog = await graphClient.GetOneDriveItemAsync(
                adminSite.Id, $"root:/{flightLogFile}");
            logger.Info("Got flight log document");

            // Get the files folder in the team OneDrive
            var teamDrive = await graphClient.GetTeamOneDriveFolderAsync(groupId, "General");
            logger.Info("Got team OneDrive General folder");

            // Copy the file from SharePoint to team drive
            var teamDriveReference = new ItemReference
            {
                DriveId = teamDrive.ParentReference.DriveId,
                Id = teamDrive.Id
            };

            await graphClient.CopySharePointFileAsync(adminSite.Id, flightLog.Id, teamDriveReference);
            logger.Info("Copied file to team files");
        }

        private static async Task CreatePreflightPlanAsync(GraphService graphClient, string groupId, DateTimeOffset departureTime)
        {
            // Create a "Pre-flight checklist" plan
            var preFlightCheckList = new Plan
            {
                Title = "Pre-flight Checklist",
                Owner = groupId
            };

            var createdPlan = await graphClient.CreatePlanAsync(preFlightCheckList);
            logger.Info("Create plan");

            // Create buckets
            var toDoBucket = new Bucket
            {
                Name = "To Do",
                PlanId = createdPlan.Id
            };

            var createdToDoBucket = await graphClient.CreateBucketAsync(toDoBucket);

            var completedBucket = new Bucket
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
        }

        private static async Task CreateChallengingPassengersListAsync(GraphService graphClient, string groupId)
        {
            // Get the team site
            var teamSite = await graphClient.GetTeamSiteAsync(groupId);

            var challengingPassengers = new SharePointList
            {
                DisplayName = "Challenging Passengers",
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
            };

            // Create the list
            var createdList = await graphClient.CreateSharePointListAsync(teamSite.Id, challengingPassengers);
        }

        private static string GetTimestamp()
        {
            var now = DateTime.Now;
            return $"{now.Hour}{now.Minute}{now.Second}";
        }
    }
}
