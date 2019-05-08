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
        [JsonProperty(PropertyName = "Description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "FlightNumber")]
        public float FlightNumber { get; set; }

        public List<LookupField> Pilots { get; set; }

        [JsonProperty(PropertyName = "FlightAttendants")]
        public List<LookupField> FlightAttendants { get; set; }

        [JsonProperty(PropertyName = "CateringLiaison")]
        public string CateringLiaison { get; set; }

        [JsonProperty(PropertyName = "DepartureTime")]
        public DateTime DepartureTime { get; set; }

        [JsonProperty(PropertyName = "DepartureGate")]
        public string DepartureGate { get; set; }
    }
}
