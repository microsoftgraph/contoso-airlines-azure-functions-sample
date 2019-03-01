// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.v
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;

namespace CreateFlightTeam.Graph
{

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class GraphResource
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string Id { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Group : GraphResource
    {
        public const string SchemaExtensionName = "ext8giz9c7n_contosoFlightTeam";

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

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ProvisioningExtension
    {
        public int SharePointItemId { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Team
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public TeamGuestSettings GuestSettings { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class TeamGuestSettings
    {
        public bool AllowCreateUpdateChannels { get; set; }
        public bool AllowDeleteChannels { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class User : GraphResource
    {
        public string DisplayName { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Invitation
    {
        public string InvitedUserEmailAddress { get; set; }
        public string InviteRedirectUrl { get; set; }
        public bool SendInvitationMessage { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public User InvitedUser { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class AddUserToGroup
    {
        [JsonProperty(PropertyName = "@odata.id")]
        public string UserPath { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Channel : GraphResource
    {
        public string DisplayName { get; set; }
        public string Description { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ChatMessage
    {
        public ItemBody Body { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ItemBody
    {
        public string Content { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class TeamsApp
    {
        [JsonIgnore]
        public string AppId { get; set; }
        [JsonProperty(PropertyName = "teamsApp@odata.bind")]
        public string BoundAppId
        {
            get
            {
                return $"https://graph.microsoft.com/beta/appCatalogs/teamsApps/{AppId}";
            }
        }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Site : GraphResource
    {
        public SiteCollection SiteCollection { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class SiteCollection
    {
        public string Hostname { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Drive : GraphResource
    {
        public string Name { get; set; }
        public string DriveType { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class DriveItem : GraphResource
    {
        public string ETag { get; set; }
        public ItemReference ParentReference { get; set; }
        public FileFacet File { get; set; }
        public DeletedFacet Deleted { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class FileFacet
    {
        public string MimeType { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class DeletedFacet
    {
        public string State { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ItemReference
    {
        public string DriveId { get; set; }
        public string Id { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Plan : GraphResource
    {
        public string Title { get; set; }
        public string Owner { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Bucket : GraphResource
    {
        public string Name { get; set; }
        public string PlanId { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class PlannerTask : GraphResource
    {
        public string Title { get; set; }
        public string PlanId { get; set; }
        public string BucketId { get; set; }
        public DateTimeOffset DueDateTime { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class SharePointList : GraphResource
    {
        public string DisplayName { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public List<ColumnDefinition> Columns { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ColumnDefinition
    {
        public string Name { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public TextColumn Text { get; set; }

    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class TextColumn
    {
        [JsonProperty(PropertyName = "@odata.type")]
        public string Type { get { return "microsoft.graph.textColumn"; } }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
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

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class NotificationPayload
    {
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string RawContent { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public NotificationVisualContent VisualContent { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class NotificationVisualContent
    {
        public string Title { get; set; }
        public string Body { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class NotificationTargetPolicy
    {
        public string[] PlatformTypes { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class TeamsChannelTab : GraphResource
    {
        public string Name { get; set; }
        public string TeamsAppId { get; set; }
        public TeamsChannelTabConfiguration Configuration { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class TeamsChannelTabConfiguration
    {
        public string EntityId { get; set; }
        public string ContentUrl { get; set; }
        public string RemoveUrl { get; set; }
        public string WebsiteUrl { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class SharePointPage : GraphResource
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public List<SharePointWebPart> WebParts { get; set; }
        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string WebUrl { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Subscription : GraphResource
    {
        public string Resource { get; set; }
        public string ChangeType { get; set; }
        public string ClientState { get; set; }
        public string NotificationUrl { get; set; }
        public DateTime ExpirationDateTime { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class SharePointWebPart
    {
        public const string ListWebPart = "f92bf067-bc19-489e-a556-7fe95f508720";

        public string Type { get; set; }
        public WebPartData Data { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class WebPartData
    {
        public string DataVersion { get; set; }
        public object Properties { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ListProperties
    {
        public bool IsDocumentLibrary { get; set; }
        public string SelectedListId { get; set; }
        public int WebpartHeightKey { get; set; }
    }

    public class LookupField
    {
        [JsonProperty(PropertyName = "LookupValue")]
        public string DisplayName { get; set; }
        public string Email { get; set; }
    }

    public class ListFields
    {
        [JsonProperty(PropertyName = "Description0")]
        public string Description { get; set; }
        [JsonProperty(PropertyName = "Flight_x0020_Number")]
        public float FlightNumber { get; set; }
        public List<LookupField> Pilots { get; set; }
        public List<LookupField> FlightAttendants { get; set; }
        public string CateringLiaison { get; set; }
        public DateTime DepartureTime { get; set; }
        public string DepartureGate { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class Identity : GraphResource
    {
        public string DisplayName { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class IdentitySet
    {
        public Identity User { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ListItem : GraphResource
    {
        public IdentitySet CreatedBy { get; set; }
        public ListFields Fields { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class GraphCollection<T>
    {
        public List<T> Value { get; set; }
        [JsonProperty(PropertyName = "@odata.nextLink")]
        public string NextLink { get; set; }
        [JsonProperty(PropertyName = "@odata.deltaLink")]
        public string DeltaLink { get; set; }
    }
}
