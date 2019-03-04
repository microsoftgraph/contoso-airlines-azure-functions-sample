using CreateFlightTeam.Graph;
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
            if (string.IsNullOrEmpty(listItem.Fields.Description) ||
                string.IsNullOrEmpty(listItem.Fields.DepartureGate) ||
                listItem.Fields.FlightNumber <= 0 ||
                listItem.Fields.DepartureTime == DateTime.MinValue)
            {
                return null;
            }

            var team = new FlightTeam();

            team.SharePointListItemId = itemId;
            team.Description = listItem.Fields.Description;
            team.FlightNumber = (int)listItem.Fields.FlightNumber;
            team.DepartureGate = listItem.Fields.DepartureGate;
            team.DepartureTime = listItem.Fields.DepartureTime;
            team.Admin = listItem.CreatedBy.User.Id;
            team.CateringLiaison = listItem.Fields.CateringLiaison;

            team.Pilots = new List<string>();
            foreach (var value in listItem.Fields.Pilots)
            {
                team.Pilots.Add(value.Email);
            }

            team.FlightAttendants = new List<string>();
            foreach (var value in listItem.Fields.FlightAttendants)
            {
                team.FlightAttendants.Add(value.Email);
            }

            return team;
        }
    }
}
