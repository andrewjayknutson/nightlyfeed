using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightlyRouteToSlack.Utilities
{
    public class ShackWagesMessageModel
    {
        public string LastDate { get; set; }
        public string Revenue { get; set; }
        public float Wages { get; set; }

        public string NetWagePercentage { get; set; }
        public string GoalWagePercentage { get; set; }

        public string OHWagePercentage { get; set; }
        public string OHGoalWagePercentage { get; set; }

        public string MktgWagePercentage { get; set; }
        public string MktgGoalWagePercentage { get; set; }

        public string SalesWagePercentage { get; set; }
        public string SalesGoalWagePercentage { get; set; }

        public string AdminWagePercentage { get; set; }
        public string AdminGoalWagePercentage { get; set; }

        public string OverallOHPercentage { get; set; }
        public string OverallOHGoalPercentage { get; set; }

        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string SalesPace { get; set; }
        public string ProductionPace { get; set; }
        public string OverallRevenueMTD { get; set; }

        public string LightsRevenue { get; set; }
        public string LightsRevenuePace { get; set; }

        public string LightsSales { get; set; }
        public string LightsSalesPace { get; set; }

        public string DetailingRevenue { get; set; }
        public string DetailingRevenuePace { get; set; }

        public string DetailingSales { get; set; }
        public string DetailingSalesPace { get; set; }
    }
}
