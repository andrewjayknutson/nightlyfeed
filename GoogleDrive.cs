using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Utilities360Wow
{
    public class GoogleDriveFile
    {
        public string Kind { get; set; }
        public string ID { get; set; }
        public string Name { get; set; }
        public string MimeType { get; set; }
    }


    public class GoogleDrive
    {

        //public string UploadFile(string accessToken)
        //{
        //    ////retrieve google API token
        //    //string ga = "";
        //    //ga = GoogleDriveAPIToken();


        //    //I have a token ... now upload the file
        //    string boundary = Guid.NewGuid().ToString().Trim().Replace("-", "");

        //    WebRequest request = WebRequest.Create("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart");
        //    request.Method = "POST";
        //    request.ContentType = "multipart/related; boundary=" + boundary;
        //    request.Headers["Authorization"] = "Bearer " + accessToken;

        //    //text body 
        //    string bodyString = "--" + boundary + Environment.NewLine + "Content-Type: application/json; charset=UTF-8" + Environment.NewLine + "{ \"name\": \"square_upload.iif\"," + Environment.NewLine + "\"parents\": [\"1sVo-H4jan6faLQv5uGeWSNtiVm381AHF\"]}" + Environment.NewLine + Environment.NewLine;
        //    byte[] bodyStringBytes = Encoding.UTF8.GetBytes(bodyString);

        //    //image body 
        //    string imageString = "--" + boundary + Environment.NewLine + "Content-Type: text/plain" + Environment.NewLine + Environment.NewLine;
        //    byte[] imageStringBytes = Encoding.UTF8.GetBytes(imageString);

        //    //actual image
        //    string fileName = @"c:\temp\testing.txt";
        //    FileStream oFileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        //    byte[] FileByteArrayData = new byte[oFileStream.Length];
        //    oFileStream.Read(FileByteArrayData, 0, System.Convert.ToInt32(oFileStream.Length));
        //    oFileStream.Close();

        //    //ending boundary
        //    string endingString = Environment.NewLine + Environment.NewLine + "--" + boundary + "--";
        //    byte[] endingStringBytes = Encoding.UTF8.GetBytes(endingString);

        //    //total length of request
        //    request.ContentLength = bodyStringBytes.Length + imageStringBytes.Length + FileByteArrayData.Length + endingStringBytes.Length;

        //    Stream dataStream = request.GetRequestStream();
        //    dataStream.Write(bodyStringBytes, 0, bodyStringBytes.Length);
        //    dataStream.Write(imageStringBytes, 0, imageStringBytes.Length);
        //    dataStream.Write(FileByteArrayData, 0, FileByteArrayData.Length);
        //    dataStream.Write(endingStringBytes, 0, endingStringBytes.Length);
        //    dataStream.Close();

        //    WebResponse response = request.GetResponse();
        //    dataStream = response.GetResponseStream();
        //    StreamReader reader = new StreamReader(dataStream);
        //    string responseFromServer = reader.ReadToEnd();

        //    reader.Close();
        //    dataStream.Close();
        //    response.Close();




        //    //string boundary = Guid.NewGuid().ToString().Trim().Replace("-", "");

        //    //WebRequest request = WebRequest.Create("https://www.googleapis.com/upload/drive/v3/files?uploadType=media");
        //    //request.Method = "POST";
        //    //request.ContentType = "text/plain";
        //    //request.Headers["Authorization"] = "Bearer " + accessToken;

        //    //////text body 
        //    ////string bodyString = "{ Name='gj_square_download',Parents = [{'1sVo-H4jan6faLQv5uGeWSNtiVm381AHF'}]}";
        //    ////byte[] bodyStringBytes = Encoding.UTF8.GetBytes(bodyString);

        //    //////image body 
        //    ////string imageString = "--" + boundary + Environment.NewLine + "Content-Disposition: form-data; name=\"image.jpg\"; filename=\"image.jpg\"" + Environment.NewLine + "Content-Type: image/jpeg " + Environment.NewLine + Environment.NewLine;
        //    ////byte[] imageStringBytes = Encoding.UTF8.GetBytes(imageString);

        //    ////actual file
        //    //string fileName = @"c:\temp\testing.txt";
        //    //FileStream oFileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        //    //byte[] FileByteArrayData = new byte[oFileStream.Length];
        //    //oFileStream.Read(FileByteArrayData, 0, System.Convert.ToInt32(oFileStream.Length));
        //    //oFileStream.Close();

        //    ////ending boundary
        //    ////string endingString = Environment.NewLine + Environment.NewLine + "--" + boundary + "--";
        //    ////byte[] endingStringBytes = Encoding.UTF8.GetBytes(endingString);

        //    ////total length of request
        //    ////request.ContentLength = bodyStringBytes.Length + imageStringBytes.Length + FileByteArrayData.Length + endingStringBytes.Length;
        //    //request.ContentLength = FileByteArrayData.Length;

        //    //Stream dataStream = request.GetRequestStream();
        //    ////dataStream.Write(bodyStringBytes, 0, bodyStringBytes.Length);
        //    ////dataStream.Write(imageStringBytes, 0, imageStringBytes.Length);
        //    //dataStream.Write(FileByteArrayData, 0, FileByteArrayData.Length);
        //    ////dataStream.Write(endingStringBytes, 0, endingStringBytes.Length);
        //    //dataStream.Close();

        //    //WebResponse response = request.GetResponse();
        //    //dataStream = response.GetResponseStream();
        //    //StreamReader reader = new StreamReader(dataStream);
        //    //string responseFromServer = reader.ReadToEnd();

        //    //reader.Close();
        //    //dataStream.Close();
        //    //response.Close();


        //    GoogleDriveFile gdf = JsonConvert.DeserializeObject<GoogleDriveFile>(responseFromServer);

        //    return gdf.ID;
        //}


        //public bool AssignFileParent(string fileID, string accessToken)
        //{
        //    //get the ID from response ... update MetaData
        //    WebRequest request = WebRequest.Create("https://www.googleapis.com/drive/v3/files/" + fileID + "?addParents=1sVo-H4jan6faLQv5uGeWSNtiVm381AHF");
        //    request.Method = "PATCH";
        //    request.Headers["Authorization"] = "Bearer " + accessToken;

        //    ////text body 
        //    string bodyString = "{ \"name\": \"square_upload_149.iif\" }";
        //    byte[] bodyStringBytes = Encoding.UTF8.GetBytes(bodyString);

        //    ////image body 
        //    //string imageString = "--" + boundary + Environment.NewLine + "Content-Disposition: form-data; name=\"image.jpg\"; filename=\"image.jpg\"" + Environment.NewLine + "Content-Type: image/jpeg " + Environment.NewLine + Environment.NewLine;
        //    //byte[] imageStringBytes = Encoding.UTF8.GetBytes(imageString);

        //    //actual file
        //    //fileName = @"c:\temp\gj_square_download_1_10_2019_40485.1147648.iif";
        //    //FileStream oFileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        //    //byte[] FileByteArrayData = new byte[oFileStream.Length];
        //    //oFileStream.Read(FileByteArrayData, 0, System.Convert.ToInt32(oFileStream.Length));
        //    //oFileStream.Close();

        //    //ending boundary
        //    //string endingString = Environment.NewLine + Environment.NewLine + "--" + boundary + "--";
        //    //byte[] endingStringBytes = Encoding.UTF8.GetBytes(endingString);

        //    //total length of request
        //    //request.ContentLength = bodyStringBytes.Length + imageStringBytes.Length + FileByteArrayData.Length + endingStringBytes.Length;
        //    request.ContentLength = bodyStringBytes.Length;

        //    Stream dataStream = request.GetRequestStream();
        //    dataStream.Write(bodyStringBytes, 0, bodyStringBytes.Length);
        //    //dataStream.Write(imageStringBytes, 0, imageStringBytes.Length);
        //    //dataStream.Write(FileByteArrayData, 0, FileByteArrayData.Length);
        //    //dataStream.Write(endingStringBytes, 0, endingStringBytes.Length);
        //    dataStream.Close();

        //    WebResponse response = request.GetResponse();
        //    dataStream = response.GetResponseStream();
        //    StreamReader reader = new StreamReader(dataStream);
        //    string responseFromServer = reader.ReadToEnd();

        //    reader.Close();
        //    dataStream.Close();
        //    response.Close();

        //    return true;

        //}



        //public string GoogleDriveAPIToken()
        //{
        //    string accessToken = "";

        //    try
        //    {
        //        string rcData = "";

        //        WebRequest request = WebRequest.Create("https://www.googleapis.com/oauth2/v4/token");
        //        request.Method = "POST";

        //        //how to get the refresh token:   https://developers.google.com/oauthplayground
        //        //drive upload client id:  670018949823-lcb2orqeea1afgv9ovc9btf9v9dcb4a5.apps.googleusercontent.com
        //        //drive upload client secret: f7Yq_QFryiJoCoMBY7Xjhfax
        //        //refresh token:  1/gYG8kbs8pqj1chN5nbFly5pbVoisSV0xekR7vGPgQzE


        //        rcData = "client_secret=f7Yq_QFryiJoCoMBY7Xjhfax&grant_type=refresh_token&refresh_token=1/gYG8kbs8pqj1chN5nbFly5pbVoisSV0xekR7vGPgQzE&client_id=670018949823-lcb2orqeea1afgv9ovc9btf9v9dcb4a5.apps.googleusercontent.com";
        //        string postData = rcData;
        //        byte[] byteArray = Encoding.UTF8.GetBytes(postData);

        //        request.ContentType = "application/x-www-form-urlencoded";
        //        request.ContentLength = byteArray.Length;

        //        Stream dataStream = request.GetRequestStream();
        //        dataStream.Write(byteArray, 0, byteArray.Length);
        //        dataStream.Close();

        //        WebResponse response = request.GetResponse();

        //        dataStream = response.GetResponseStream();
        //        StreamReader reader = new StreamReader(dataStream);
        //        string responseFromServer = reader.ReadToEnd();

        //        reader.Close();
        //        dataStream.Close();
        //        response.Close();

        //        int firstColon = responseFromServer.IndexOf(":");
        //        int firstComma = responseFromServer.IndexOf(",");

        //        accessToken = responseFromServer.Substring(firstColon + 3, firstComma - (firstColon + 4));

        //    }
        //    catch (Exception e)
        //    {
        //        //send this error to Slack

        //        accessToken = "ERROR:  GoogleAPIToken Error";
        //    }

        //    return accessToken;


        //}


    }
}
