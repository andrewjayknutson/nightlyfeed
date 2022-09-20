using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightlyRouteToSlack.Reviews
{
    public class ReviewsRoot
    {
        public List<Review> reviews { get; set; }
        public int totalReviewCount { get; set; }
    }
}
