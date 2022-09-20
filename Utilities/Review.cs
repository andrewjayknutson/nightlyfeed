using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Net;
using NightlyRouteToSlack.Reviews;
using System.Runtime.Serialization.Json;
using Utilities360Wow;

namespace NightlyRouteToSlack.Utilities
{
    class Review
    {
        public string Reviewer { get; set; }
        public string ReviewDate { get; set; }
        public string ReviewStars { get; set; }
        public string ReviewReview { get; set; }
        public string ReviewReplyDate { get; set; }
        public string ReviewReplyReview { get; set; }


        public int ReturnReviewsNumber(string bizUnit, string currentMonth, string currentYear)
        {
            try
            {
                ConfigSettings cs = new ConfigSettings();
                Utilities360Wow.Google g = new Utilities360Wow.Google();

                string accountsID = cs.ReturnConfigSetting("NightlyRouteToSlack", "GoogleAPIsAccountID");
                int totalReviews = 0;

                //get the token
                string ga = "";
                ga = g.GetGoogleAPIToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "GoogleClientSecret"), cs.ReturnConfigSetting("NightlyRouteToSlack", "GoogleRefreshToken"));

                //then call this web request
                string reviewsID = "";
                switch (bizUnit)
                {
                    case "gj":
                        reviewsID = cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkReviewsID");
                        break;
                    case "ym":
                        reviewsID = cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveReviewsID");
                        break;
                    case "ss":
                        reviewsID = cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackReviewsID");
                        break;
                    default:
                        reviewsID = cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkReviewsID");
                        break;
                }
                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/" + accountsID + "/locations/" + reviewsID + "/reviews");
                webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + ga);
                webReq.Method = "GET";
                HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                Stream answer = webResp.GetResponseStream();
                StreamReader _answer = new StreamReader(answer);


                ReviewsRoot deserializedReview = new ReviewsRoot();
                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
                DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedReview.GetType());
                deserializedReview = ser.ReadObject(ms) as ReviewsRoot;
                ms.Close();


                Review r = new Review();
                List<Review> localList = new List<Review>();
                try
                {
                    for (int i = 0; i < deserializedReview.totalReviewCount; i++)
                    {
                        if (currentMonth == Convert.ToDateTime(DateTime.Parse(deserializedReview.reviews[i].createTime).ToShortDateString()).Month.ToString() && currentYear == Convert.ToDateTime(DateTime.Parse(deserializedReview.reviews[i].createTime).ToShortDateString()).Year.ToString())
                        {
                            totalReviews = totalReviews + 1;
                        }

                    }
                }
                catch
                {
                }

                return totalReviews;

            }
            catch
            {
                return 0;
            }

        }


        public double ReturnReviewsRequested(string bizUnit)
        {
            CommonClient cc = new CommonClient();
            //pull review requests sent for the month and year and this biz

            double totalSent = 0;
            dbConnect dc = new dbConnect();
            dc.OpenMessageConnection();

            //**** Total Revenue
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetSentMessageByDate";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = cc.GetStartDate();
            //cmd.Parameters["@startDate"].Value = "1/1/2019";

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1);
            //cmd.Parameters["@endDate"].Value = "3/31/2019";

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = bizUnit;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                totalSent = Convert.ToDouble(dt.Rows[0]["bizUnitCount"].ToString());
            }

            return totalSent;

        }


    }
}
