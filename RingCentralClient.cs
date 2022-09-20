using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;

namespace Utilities360Wow
{
    public class RingCentralClient
    {

        public string GetRingCentralToken(string ringCentralTokenURL, string ringCentralTokenData, string ringCentralTokenAuth)
        {
            WebRequest request = WebRequest.Create(ringCentralTokenURL);
            request.Method = "POST";

            byte[] byteArray = Encoding.UTF8.GetBytes(ringCentralTokenData);

            request.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
            request.ContentLength = byteArray.Length;
            request.Headers["Authorization"] = ringCentralTokenAuth;

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

            return responseFromServer.Substring(firstColon + 3, firstComma - (firstColon + 4));

        }


        public void SendSMSRingCentral(string accessToken, string ringCentralSMSURL, string ringCentralSMSPhone, string Phone, string Message)
        {
            WebRequest request = WebRequest.Create(ringCentralSMSURL);
            request.Method = "POST";

            byte[] byteArray = Encoding.UTF8.GetBytes("{\"to\": [{\"phoneNumber\": \"+" + Phone + "\"}],\"from\": {\"phoneNumber\": \"+" + ringCentralSMSPhone + "}\"},\"text\": \"" + Message + "\"}");

            request.ContentType = "application/json";
            request.ContentLength = byteArray.Length;
            request.Headers["Authorization"] = "Bearer " + accessToken;

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



        public void SendMMSRingCentral(string accessToken, string ringCentralMMSURL, string ringCentralMMSPhone, string ringCentralMMSImageLocation, string Phone, string Message, string ImageName)
        {
            string boundary = Guid.NewGuid().ToString().Trim().Replace("-", "");

            WebRequest request = WebRequest.Create(ringCentralMMSURL);
            request.Method = "POST";
            request.ContentType = "multipart/mixed; boundary=" + boundary;
            request.Headers["Authorization"] = "Bearer " + accessToken;

            //text body 
            string bodyString = "--" + boundary + Environment.NewLine + "Content-Type: application/json; charset=UTF-8" + Environment.NewLine + "Content-Transfer-Encoding: 8bit" + Environment.NewLine + Environment.NewLine + "{\"to\" :[{\"phoneNumber\": \"+" + Phone + "\"}]," + Environment.NewLine + "\"text\" :\"" + Message + "\" ," + Environment.NewLine;
            bodyString = bodyString + "\"from\" :{\"phoneNumber\": \"+" + ringCentralMMSPhone + "\"}}" + Environment.NewLine + Environment.NewLine;

            byte[] bodyStringBytes = Encoding.UTF8.GetBytes(bodyString);

            //image body 
            string imageString = "--" + boundary + Environment.NewLine + "Content-Disposition: form-data; name=\"image.jpg\"; filename=\"image.jpg\"" + Environment.NewLine + "Content-Type: image/jpeg " + Environment.NewLine + Environment.NewLine;
            byte[] imageStringBytes = Encoding.UTF8.GetBytes(imageString);

            //actual image
            string fileName = ringCentralMMSImageLocation + ImageName;
            FileStream oFileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            byte[] FileByteArrayData = new byte[oFileStream.Length];
            oFileStream.Read(FileByteArrayData, 0, System.Convert.ToInt32(oFileStream.Length));
            oFileStream.Close();

            //ending boundary
            string endingString = Environment.NewLine + Environment.NewLine + "--" + boundary + "--";
            byte[] endingStringBytes = Encoding.UTF8.GetBytes(endingString);

            //total length of request
            request.ContentLength = bodyStringBytes.Length + imageStringBytes.Length + FileByteArrayData.Length + endingStringBytes.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(bodyStringBytes, 0, bodyStringBytes.Length);
            dataStream.Write(imageStringBytes, 0, imageStringBytes.Length);
            dataStream.Write(FileByteArrayData, 0, FileByteArrayData.Length);
            dataStream.Write(endingStringBytes, 0, endingStringBytes.Length);
            dataStream.Close();

            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseFromServer = reader.ReadToEnd();

            reader.Close();
            dataStream.Close();
            response.Close();

        }



    }
}
