// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using Newtonsoft.Json;
using System;

namespace create_flight_team.Models
{
    class CreateFlightTeamResponse
    {
        public string Result { get; set; }

        [JsonProperty(DefaultValueHandling = DefaultValueHandling.Ignore)]
        public Exception Details { get; set; }
    }
}
