using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;

namespace Utilities360Wow
{
    public class iAuditorClient
    {

        public string GetiAuditorAPIToken(string grantCredentials)
        {
            WebRequest request = WebRequest.Create("https://api.safetyculture.io/auth");
            request.Method = "POST";

            byte[] byteArray = Encoding.UTF8.GetBytes(grantCredentials);

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

            return responseFromServer.Substring(firstColon + 2, firstComma - (firstColon + 3));

        }

        public string GetMoveiAuditorAPIToken(string grantCredentials)
        {
            WebRequest request = WebRequest.Create("https://api.safetyculture.io/auth");
            request.Method = "POST";

            byte[] byteArray = Encoding.UTF8.GetBytes(grantCredentials);

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

            return responseFromServer.Substring(firstColon + 2, firstComma - (firstColon + 3));

        }

    }
}
