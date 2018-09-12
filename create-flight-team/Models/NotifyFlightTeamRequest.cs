using System;
using System.Collections.Generic;
using System.Text;

namespace create_flight_team.Models
{
    class NotifyFlightTeamRequest
    {
        public float FlightNumber { get; set; }
        public string NewDepartureGate { get; set; }
    }
}
