using CreateFlightTeam.Graph;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;

namespace CreateFlightTeam.Models
{
    public class FlightTeam
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        [JsonProperty(PropertyName = "sharePointListItemId")]
        public string SharePointListItemId { get; set; }
        [JsonProperty(PropertyName = "teamId")]
        public string TeamId { get; set; }
        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }
        [JsonProperty(PropertyName = "flightNumber")]
        public int FlightNumber { get; set; }
        [JsonProperty(PropertyName = "departureGate")]
        public string DepartureGate { get; set; }
        [JsonProperty(PropertyName = "departureTime")]
        public DateTime DepartureTime { get; set; }
        [JsonProperty(PropertyName = "admin")]
        public string Admin { get; set; }
        [JsonProperty(PropertyName = "pilots")]
        public List<string> Pilots { get; set; }
        [JsonProperty(PropertyName = "flightAttendants")]
        public List<string> FlightAttendants { get; set; }
        [JsonProperty(PropertyName = "cateringLiaison")]
        public string CateringLiaison { get; set; }

        public FlightTeam() { }
        public static FlightTeam FromListItem(string itemId, ListItem listItem)
        {
            var jsonFields = JsonConvert.SerializeObject(listItem.Fields.AdditionalData);
            var fields = JsonConvert.DeserializeObject<ListFields>(jsonFields);

            if (string.IsNullOrEmpty(fields.Description) ||
                string.IsNullOrEmpty(fields.DepartureGate) ||
                fields.FlightNumber <= 0 ||
                fields.DepartureTime == DateTime.MinValue)
            {
                return null;
            }

            var team = new FlightTeam();

            team.SharePointListItemId = itemId;
            team.Description = fields.Description;
            team.FlightNumber = (int)fields.FlightNumber;
            team.DepartureGate = fields.DepartureGate;
            team.DepartureTime = fields.DepartureTime;
            team.Admin = listItem.CreatedBy.User.Id;
            team.CateringLiaison = fields.CateringLiaison;

            team.Pilots = new List<string>();
            if (fields.Pilots != null) {
                foreach (var value in fields.Pilots)
                {
                    team.Pilots.Add(value.Email);
                }
            }

            team.FlightAttendants = new List<string>();
            if (fields.FlightAttendants != null) {
                foreach (var value in fields.FlightAttendants)
                {
                    team.FlightAttendants.Add(value.Email);
                }
            }

            return team;
        }
    }
}
