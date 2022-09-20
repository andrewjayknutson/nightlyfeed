using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightlyRouteToSlack.Utilities
{
    public class MoveWagesMessageModel
    {
        public string LastDate { get; set; }
        public string Revenue { get; set; }
        public float Wages { get; set; }

        public string NetWagePercentage { get; set; }
        public string GoalWagePercentage { get; set; }
        public string OHWagePercentage { get; set; }
        public string OHGoalWagePercentage { get; set; }

        public string StartDate { get; set; }
        public string EndDate { get; set; }

        public string RevenuePace { get; set; }

        public string PopinsDone { get; set; }
        public string PopinsJobs { get; set; }
        public string PopinPercentage { get; set; }
        public string PopinGoalPercentage { get; set; }

        public string DamagePercentage { get; set; }
        public string DamagePercentageGoal { get; set; }
        public string Damage { get; set; }
        public string DamageGoal { get; set; }
        public string ServicePercentage { get; set; }
        public string ServicePercentageGoal { get; set; }
        public string Service { get; set; }
        public string ServiceGoal { get; set; }

        public string GrossOverheadPercentage { get; set; }
        public string NetOverheadPercentage { get; set; }
        public string OverheadGoal { get; set; }
        public string OverheadGoalResults { get; set; }


    }
}
