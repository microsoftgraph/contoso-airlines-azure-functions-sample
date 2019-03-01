using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;

namespace CreateFlightTeam.Models
{
    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public class ChangeNotification
    {
        public string SubscriptionId { get; set; }
        public DateTime SubscriptionExpirationDateTime { get; set; }
        public string ClientState { get; set; }
        public string ChangeType { get; set; }
        public string Resource { get; set; }
    }

    public class ListChangedRequest
    {
        [JsonProperty("value")]
        public List<ChangeNotification> Changes { get; set; }
    }
}
