using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightlyRouteToSlack.Utilities
{
    public class MoveDailyWages
    {

        public string startDate { get; set; }
        public string endDate { get; set; }
        public string thruDate { get; set; }
        public string lastUpdated { get; set; }
        public string score { get; set; }
        public string goal { get; set; }

        public string avg { get; set; }

        public List<MoveDailyWage> DailyWages { get; set; }
    }
}
