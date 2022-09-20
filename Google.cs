using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using Utilities360Wow.Reviews;

namespace Utilities360Wow
{
    public class Google
    {
        public List<GoogleReview> GetReviewListByDateRange(string googleAPIToken, string reviewsID, string accountID, string startDate, string endDate)
        {
            int howMany = 0;
            string pageToken = "";
            List<GoogleReview> localList = new List<GoogleReview>();
            ReviewsRoot deserializedReview = new ReviewsRoot();

            try
            {
                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/" + accountID + "/locations/" + reviewsID + "/reviews");
                webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + googleAPIToken);
                webReq.Method = "GET";
                HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                Stream answer = webResp.GetResponseStream();
                StreamReader _answer = new StreamReader(answer);


                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
                DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedReview.GetType());
                deserializedReview = ser.ReadObject(ms) as ReviewsRoot;
                ms.Close();

            }
            catch (Exception e)
            {
                //returns nothing
                List<GoogleReview> catchList = new List<GoogleReview>();
                return catchList;
            }


            //grab the pageToken
            pageToken = deserializedReview.nextPageToken;

            //read into the localList the first set of reviews
            try
            {
                GoogleReview r = new GoogleReview();
                for (int i = 0; i < deserializedReview.reviews.Count; i++)
                {
                    if (DateTime.Parse(deserializedReview.reviews[i].createTime.ToString()) > DateTime.Parse(startDate) & DateTime.Parse(deserializedReview.reviews[i].createTime.ToString()) < DateTime.Parse(endDate))
                    {
                        r = new GoogleReview();
                        r.Reviewer = deserializedReview.reviews[i].reviewer.displayName;
                        r.ReviewDate = DateTime.Parse(deserializedReview.reviews[i].createTime).ToShortDateString();
                        switch (deserializedReview.reviews[i].starRating)
                        {
                            case "FIVE":
                                r.ReviewStars = "5star.png";
                                break;
                            case "FOUR":
                                r.ReviewStars = "4star.png";
                                break;
                            case "THREE":
                                r.ReviewStars = "3star.png";
                                break;
                            case "TWO":
                                r.ReviewStars = "2star.png";
                                break;
                            case "ONE":
                                r.ReviewStars = "1star.png";
                                break;
                            default:
                                r.ReviewStars = "5star.png";
                                break;

                        }
                        r.ReviewReview = deserializedReview.reviews[i].comment;

                        if (deserializedReview.reviews[i].reviewReply != null)
                        {
                            r.ReviewReplyDate = DateTime.Parse(deserializedReview.reviews[i].reviewReply.updateTime).ToShortDateString();
                            r.ReviewReplyReview = deserializedReview.reviews[i].reviewReply.comment;
                        }

                        localList.Add(r);
                    }
                }
            }
            catch (Exception e)
            {
                //send this error to Slack

                List<GoogleReview> catchList = new List<GoogleReview>();
                return catchList;

            }

            //counting how many times we loop through the paginations
            howMany = howMany + 1;


            while (!(pageToken is null))
            {
                //grab the next set of reviews
                try
                {
                    HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/" + accountID + "/locations/" + reviewsID + "/reviews?pageToken=" + pageToken);
                    webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + googleAPIToken);
                    webReq.Method = "GET";
                    HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                    Stream answer = webResp.GetResponseStream();
                    StreamReader _answer = new StreamReader(answer);


                    MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedReview.GetType());
                    deserializedReview = ser.ReadObject(ms) as ReviewsRoot;
                    ms.Close();

                }
                catch (Exception e)
                {
                    //returns nothing
                    List<GoogleReview> catchList = new List<GoogleReview>();
                    return catchList;
                }

                //grab the pageToken
                pageToken = deserializedReview.nextPageToken;
                howMany = howMany + 1;

                //read into the localList the first set of reviews
                try
                {
                    GoogleReview r = new GoogleReview();
                    for (int i = 0; i < deserializedReview.reviews.Count; i++)
                    {
                        if (DateTime.Parse(deserializedReview.reviews[i].createTime.ToString()) > DateTime.Parse(startDate) & DateTime.Parse(deserializedReview.reviews[i].createTime.ToString()) < DateTime.Parse(endDate))
                        {
                            r = new GoogleReview();
                            r.Reviewer = deserializedReview.reviews[i].reviewer.displayName;
                            r.ReviewDate = DateTime.Parse(deserializedReview.reviews[i].createTime).ToShortDateString();
                            switch (deserializedReview.reviews[i].starRating)
                            {
                                case "FIVE":
                                    r.ReviewStars = "5star.png";
                                    break;
                                case "FOUR":
                                    r.ReviewStars = "4star.png";
                                    break;
                                case "THREE":
                                    r.ReviewStars = "3star.png";
                                    break;
                                case "TWO":
                                    r.ReviewStars = "2star.png";
                                    break;
                                case "ONE":
                                    r.ReviewStars = "1star.png";
                                    break;
                                default:
                                    r.ReviewStars = "5star.png";
                                    break;

                            }
                            r.ReviewReview = deserializedReview.reviews[i].comment;

                            if (deserializedReview.reviews[i].reviewReply != null)
                            {
                                r.ReviewReplyDate = DateTime.Parse(deserializedReview.reviews[i].reviewReply.updateTime).ToShortDateString();
                                r.ReviewReplyReview = deserializedReview.reviews[i].reviewReply.comment;
                            }

                            localList.Add(r);
                        }
                    }
                }
                catch (Exception e)
                {
                    //send this error to Slack

                    List<GoogleReview> catchList = new List<GoogleReview>();
                    return catchList;

                }

            }






            return localList;
        }


        public List<GoogleReview> GetReviewsList(string reviewsID, string clientSecret, string refreshToken, string accountID)
        {

            List<GoogleReview> localList = new List<GoogleReview>();
            ReviewsRoot deserializedReview = new ReviewsRoot();

            //retrieve google API token
            string ga = "";
            ga = GetGoogleAPIToken(clientSecret, refreshToken);

            //retrieve google reviews
            try
            {
                //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/101547720956473214676/locations/" + reviewsID + "/reviews?pageSize=100");
                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/" + accountID + "/locations/" + reviewsID + "/reviews?pageSize=100");
                webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + ga);
                webReq.Method = "GET";
                HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                Stream answer = webResp.GetResponseStream();
                StreamReader _answer = new StreamReader(answer);


                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
                DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedReview.GetType());
                deserializedReview = ser.ReadObject(ms) as ReviewsRoot;
                ms.Close();

            }
            catch (Exception e)
            {
                //send this error to Slack

                List<GoogleReview> catchList = new List<GoogleReview>();
                return catchList;
            }


            try
            {
                GoogleReview r = new GoogleReview();
                for (int i = 0; i < deserializedReview.reviews.Count; i++)
                {
                    r = new GoogleReview();
                    r.Reviewer = deserializedReview.reviews[i].reviewer.displayName;
                    r.ReviewDate = DateTime.Parse(deserializedReview.reviews[i].createTime).ToShortDateString();
                    switch (deserializedReview.reviews[i].starRating)
                    {
                        case "FIVE":
                            r.ReviewStars = "5star.png";
                            break;
                        case "FOUR":
                            r.ReviewStars = "4star.png";
                            break;
                        case "THREE":
                            r.ReviewStars = "3star.png";
                            break;
                        case "TWO":
                            r.ReviewStars = "2star.png";
                            break;
                        case "ONE":
                            r.ReviewStars = "1star.png";
                            break;
                        default:
                            r.ReviewStars = "5star.png";
                            break;

                    }
                    r.ReviewReview = deserializedReview.reviews[i].comment;

                    if (deserializedReview.reviews[i].reviewReply != null)
                    {
                        r.ReviewReplyDate = DateTime.Parse(deserializedReview.reviews[i].reviewReply.updateTime).ToShortDateString();
                        r.ReviewReplyReview = deserializedReview.reviews[i].reviewReply.comment;
                    }

                    localList.Add(r);
                }

            }
            catch (Exception e)
            {
                //send this error to Slack

                List<GoogleReview> catchList = new List<GoogleReview>();
                return catchList;

            }

            return localList;
        }

        public string GetReviewsCount(string googleAPIToken, string reviewsID, string accountID)
        {
            ReviewsRoot deserializedReview = new ReviewsRoot();

            //retrieve google reviews
            try
            {
                //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/101547720956473214676/locations/" + reviewsID + "/reviews");
                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://mybusiness.googleapis.com/v4/accounts/" + accountID + "/locations/" + reviewsID + "/reviews");
                webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + googleAPIToken);
                webReq.Method = "GET";
                HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                Stream answer = webResp.GetResponseStream();
                StreamReader _answer = new StreamReader(answer);


                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
                DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedReview.GetType());
                deserializedReview = ser.ReadObject(ms) as ReviewsRoot;
                ms.Close();

            }
            catch (Exception e)
            {
                //send this error to Slack
                return "#error_getreviewscount";
            }

            return deserializedReview.totalReviewCount.ToString();

        }




        //***************************************** PRIVATE FUNCTIONS *****************************************


        public string GetGoogleAPIToken(string clientSecret, string refreshToken)
        {
            string accessToken = "";

            try
            {
                string rcData = "";

                WebRequest request = WebRequest.Create("https://www.googleapis.com/oauth2/v4/token");
                request.Method = "POST";

                //rcData = "client_secret=vDkIo0cTng_f8oElhW3FKFhd&grant_type=refresh_token&refresh_token=1%2FaFKRF9mhKiW8gfIVjMUHU13wsbHENjxuJvzyJy-2YcA&client_id=630764660080-07l3sjjj515i462acslkjvcs98b8ce19.apps.googleusercontent.com";
                rcData = "client_secret=" + clientSecret + "&grant_type=refresh_token&refresh_token=" + refreshToken;
                string postData = rcData;
                byte[] byteArray = Encoding.UTF8.GetBytes(postData);

                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;

                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();

                WebResponse response = request.GetResponse();

                dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();

                reader.Close();
                dataStream.Close();
                response.Close();

                int firstColon = responseFromServer.IndexOf(":");
                int firstComma = responseFromServer.IndexOf(",");

                accessToken = responseFromServer.Substring(firstColon + 3, firstComma - (firstColon + 4));

            }
            catch (Exception e)
            {
                //send this error to Slack

                accessToken = "ERROR:  GoogleAPIToken Error";
            }

            return accessToken;

        }






    public string GetGoogleAPITokenByClientID(string clientID, string clientSecret, string refreshToken)
        {
            string accessToken = "";

            try
            {
                string rcData = "";

                WebRequest request = WebRequest.Create("https://www.googleapis.com/oauth2/v4/token");
                request.Method = "POST";
                rcData = "client_id=" + clientID + "&client_secret=" + clientSecret + "&refresh_token=" + refreshToken + "&grant_type=refresh_token";
                string postData = rcData;
                byte[] byteArray = Encoding.UTF8.GetBytes(postData);

                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;

                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();

                WebResponse response = request.GetResponse();

                dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();

                reader.Close();
                dataStream.Close();
                response.Close();

                int firstColon = responseFromServer.IndexOf(":");
                int firstComma = responseFromServer.IndexOf(",");

                accessToken = responseFromServer.Substring(firstColon + 3, firstComma - (firstColon + 4));

            }
            catch (Exception e)
            {
                //send this error to Slack

                accessToken = "ERROR:  GoogleAPIToken Error";
            }

            return accessToken;

        }











    }
}
