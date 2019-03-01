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
        public FlightTeam(ListItem listItem)
        {
            SharePointListItemId = listItem.Id;
            Description = listItem.Fields.Description;
            FlightNumber = (int)listItem.Fields.FlightNumber;
            DepartureGate = listItem.Fields.DepartureGate;
            DepartureTime = listItem.Fields.DepartureTime;
            Admin = listItem.CreatedBy.User.Id;
            CateringLiaison = listItem.Fields.CateringLiaison;

            Pilots = new List<string>();
            foreach (var value in listItem.Fields.Pilots)
            {
                Pilots.Add(value.Email);
            }

            FlightAttendants = new List<string>();
            foreach (var value in listItem.Fields.FlightAttendants)
            {
                FlightAttendants.Add(value.Email);
            }
        }
    }
}
