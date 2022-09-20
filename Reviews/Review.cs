using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightlyRouteToSlack.Reviews
{
    public class Review
    {
        public string reviewId { get; set; }
        public Reviewer reviewer { get; set; }
        public string starRating { get; set; }
        public string comment { get; set; }
        public string createTime { get; set; }
        public string updateTime { get; set; }
        public ReviewReply reviewReply { get; set; }
    }
}
