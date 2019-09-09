using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;

namespace CreateFlightTeam.Models
{
    public class ListSubscription
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }
        [JsonProperty(PropertyName = "expiration")]
        public DateTime Expiration { get; set; }
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }
        [JsonProperty(PropertyName = "deltaLink")]
        public string DeltaLink { get; set; }

        public bool IsExpired()
        {
            return Expiration <= DateTime.UtcNow;
        }

        public bool IsExpiredOrCloseToExpired()
        {
            // If expiration is not more than 12 hours away
            return Expiration.AddHours(-12) <= DateTime.UtcNow;
        }
    }
}
