// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.v
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;

namespace CreateFlightTeam.Graph
{
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
}
