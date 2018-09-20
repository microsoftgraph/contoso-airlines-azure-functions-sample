using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace create_flight_team.Graph
{
    public class GraphResource
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string Id { get; set; }
    }

    public class Group : GraphResource
    {
        public const string SchemaExtensionName = "ext19x0ug7l_contosoFlightTeam";

        public string Description { get; set; }
        public string DisplayName { get; set; }
        public string[] GroupTypes { get; set; }
        public bool MailEnabled { get; set; }
        public string MailNickname { get; set; }
        public bool SecurityEnabled { get; set; }
        public string Visibility { get; set; }

        [JsonProperty(PropertyName = SchemaExtensionName, DefaultValueHandling = DefaultValueHandling.Ignore)]
        public ProvisioningExtension Extension { get; set; }

        [JsonProperty(PropertyName = "members@odata.bind", DefaultValueHandling = DefaultValueHandling.Ignore)]
        public List<string> Members { get; set; }

        [JsonProperty(PropertyName = "owners@odata.bind", DefaultValueHandling = DefaultValueHandling.Ignore)]
        public List<string> Owners { get; set; }
    }

    public class ProvisioningExtension
    {
        public int SharePointItemId { get; set; }
    }

    public class Team
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public TeamGuestSettings GuestSettings { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    public class TeamGuestSettings
    {
        public bool AllowCreateUpdateChannels { get; set; }
        public bool AllowDeleteChannels { get; set; }
    }

    public class User : GraphResource { }

    public class Invitation
    {
        public string InvitedUserEmailAddress { get; set; }
        public string InviteRedirectUrl { get; set; }
        public bool SendInvitationMessage { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public User InvitedUser { get; set; }
    }

    public class AddUserToGroup
    {
        [JsonProperty(PropertyName = "@odata.id")]
        public string UserPath { get; set; }
    }

    public class Channel : GraphResource
    {
        public string DisplayName { get; set; }
        public string Description { get; set; }
    }

    public class ChatThread
    {
        public ChatMessage RootMessage { get; set; }
    }

    public class ChatMessage
    {
        public ItemBody Body { get; set; }
    }

    public class ItemBody
    {
        public string Content { get; set; }
    }

    public class TeamsApp : GraphResource { }

    public class Site : GraphResource
    {
        public SiteCollection SiteCollection { get; set; }
    }

    public class SiteCollection
    {
        public string Hostname { get; set; }
    }

    public class DriveItem : GraphResource
    {
        public ItemReference ParentReference { get; set; }
    }

    public class ItemReference
    {
        public string DriveId { get; set; }
        public string Id { get; set; }
    }

    public class Plan : GraphResource
    {
        public string Title { get; set; }
        public string Owner { get; set; }
    }

    public class Bucket : GraphResource
    {
        public string Name { get; set; }
        public string PlanId { get; set; }
    }

    public class PlannerTask : GraphResource
    {
        public string Title { get; set; }
        public string PlanId { get; set; }
        public string BucketId { get; set; }
        public DateTimeOffset DueDateTime { get; set; }
    }

    public class SharePointList : GraphResource
    {
        public string DisplayName { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public List<ColumnDefinition> Columns { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    public class ColumnDefinition
    {
        public string Name { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public TextColumn Text { get; set; }

    }

    public class TextColumn
    {
        [JsonProperty(PropertyName = "@odata.type")]
        public string Type { get { return "microsoft.graph.textColumn"; } }
    }

    public class Notification : GraphResource
    {
        public string TargetHostName { get; set; }
        public string AppNotificationId { get; set; }
        public DateTimeOffset ExpirationDateTime { get; set; }
        public NotificationPayload Payload { get; set; }
        public NotificationTargetPolicy TargetPolicy { get; set; }
        public string Priority { get; set; }
        public string GroupName { get; set; }
        public int DisplayTimeToLive { get; set; }
    }

    public class NotificationPayload
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string RawContent { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public NotificationVisualContent VisualContent { get; set; }
    }

    public class NotificationVisualContent
    {
        public string Title { get; set; }
        public string Body { get; set; }
    }

    public class NotificationTargetPolicy
    {
        public string[] PlatformTypes { get; set; }
    }

    public class TeamsChannelTab : GraphResource
    {
        public string Name { get; set; }
        public string TeamsAppId { get; set; }
        public TeamsChannelTabConfiguration Configuration { get; set; }
    }

    public class TeamsChannelTabConfiguration
    {
        public string EntityId { get; set; }
        public string ContentUrl { get; set; }
        public string RemoveUrl { get; set; }
        public string WebsiteUrl { get; set; }
    }

    public class SharePointPage : GraphResource
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public List<SharePointWebPart> WebParts { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    public class SharePointWebPart
    {
        public const string ListWebPart = "f92bf067-bc19-489e-a556-7fe95f508720";

        public string Type { get; set; }
        public WebPartData Data { get; set; }
    }

    public class WebPartData
    {
        public string DataVersion { get; set; }
        public object Properties { get; set; }
    }

    public class ListProperties
    {
        public const string DocLibraryViewId = "5c8737ce-7642-483c-86f0-a1dd698f1668";
        public const string ListViewId = "b48e4b66-7e47-499e-a79a-d238da845214";

        public bool IsDocumentLibrary { get; set; }
        public string SelectedListId { get; set; }
        public int WebpartHeightKey { get; set; }
        public string SelectedViewId { get; set; }
    }

    public class GraphCollection<T>
    {
        public List<T> Value { get; set; }
    }
}
