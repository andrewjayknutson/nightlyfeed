using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace Utilities360Wow
{
    public class SlackClient
    {

        public List<SlackUser> RetrieveUsers()
        {
            SlackUser local = new SlackUser();
            List<SlackUser> localList = new List<SlackUser>();

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://slack.com/api/users.list");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer xoxb-169851555877-904447001573-ak4PI6DIW5HcKDCjt1NFU1QK");
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            contents = "[" + contents + "]";

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                foreach (JToken tokens in x.SelectToken("members"))
                {

                    if (tokens["deleted"].ToString().ToLower() == "false")
                    {

                        local = new SlackUser();

                        local.ID = tokens["id"].ToString().ToLower();
                        local.SlackName = tokens["name"].ToString().ToLower();
                        local.RealName = tokens["real_name"].ToString().ToLower();

                        localList.Add(local);
                    }

                }
            }


            return localList;
        }



        public void PostMessage(string Message, string UserID, string UserToken)
        {
            string channel = UserID.ToUpper();
            string text = HttpUtility.UrlEncode(Message);
            string username = HttpUtility.UrlEncode("GJ Operations");
            string blocks = HttpUtility.UrlEncode("[{\"type\": \"section\",\"text\": {\"type\": \"mrkdwn\",\"text\": \"" + Message + "\"}}]");
            string as_user = "False";

            WebRequest request = WebRequest.Create("https://slack.com/api/chat.postMessage?token=" + UserToken + "&channel=" + channel + "&as_user=False&username=Operations&text=" + Message + "&blocks=" + blocks);
            //WebRequest request = WebRequest.Create("https://slack.com/api/chat.postMessage?token=" + UserToken + "&channel=" + channel + "&text=" + Message + "&as_user=" + as_user + "&blocks=" + blocks);

            request.Method = "POST";

            byte[] byteArray = Encoding.UTF8.GetBytes("");

            request.ContentType = "application/json";
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


        }






        public void SendMessage(string jsonToSend, string slackChannel)
        {
            Uri uriOut = new Uri(slackChannel);

            byte[] data = Encoding.UTF8.GetBytes(jsonToSend);

            WebRequest req = WebRequest.Create(uriOut);
            req.ContentType = "application/json";
            req.Method = "POST";
            req.ContentLength = data.Length;

            Stream streamOut = req.GetRequestStream();
            streamOut.Write(data, 0, data.Length);
            streamOut.Close();

            Stream response = req.GetResponse().GetResponseStream();
            StreamReader reader = new StreamReader(response);

            string res = reader.ReadToEnd();
            reader.Close();
            response.Close();
        }

    }
}
