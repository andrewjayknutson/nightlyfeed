using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using NightlyRouteToSlack.iAuditorAudits;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Utilities360Wow;

namespace NightlyRouteToSlack.Utilities
{
    class UtilityClass
    {

        //**************************************************************************************************************************

        //This is where I will plan to make notes of what I'm working on and when
        //  4/15:   request from Jeremy to update the nightly scorecard for GJ
        //          I have currently created sproc spAutomated_GetWagesByDatesV2
        //          I have currently updated code JunkScorecardDaily
        //          both will need to be deployed once all questions are answered and coding is finished

        //**************************************************************************************************************************



        public static int NumberFromString(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        public void ReportOutCovidTracking()
        {
            DateTime temp;
            bool foundOne = false;
            string json = "{'text': '*COVID-19 Incident Tracking*\n";

            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "Incidents/Actions Revised!A1:D500";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader("1uSMqDvvxFN0jlVz1xsXrWajC9QBjfjMNAY_9q0jwtgk", range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        //if column C (Date Entered) is equal to today, then kick it out
                        DateTime.TryParse(row.ColumnValue(2).ToString().Trim(), out temp);
                        if (DateTime.Now.ToShortDateString() == temp.ToShortDateString())
                        {
                            foundOne = true;
                            json = json + "   " + row.ColumnValue(1).ToString().Trim() + " completed an entry for " + row.ColumnValue(0).ToString().Trim() + " on " + row.ColumnValue(2).ToString().Trim() + "\n";
                        }

                    }
                    catch (WebException ex)
                    {

                    }
                }

            }
            catch (WebException ex)
            {

            }

            if (!foundOne)
            {
                json = json + "   No new entries completed on " + DateTime.Now.ToShortDateString() +"\n";
            }


            json = json + "______________________________________________________\n";
            json = json + "\n   COVID-19 Incident Tracking Doc:   https://docs.google.com/spreadsheets/d/1uSMqDvvxFN0jlVz1xsXrWajC9QBjfjMNAY_9q0jwtgk'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_covidtracking").ToString().Replace("\\\\", "\\"));

        }



        public void UploadHiringUpdate()
        {
            int thirtyDay = 0;
            string json = "";
            string currentColumn = "";
            string thirtyDayColumn = "";
            json = "{'text': '*People Needs*\n";

            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            //1-800-GOT-JUNK?
            //  get current and 30 day columns
            string spreadSheetId = "1IsdgX64l-JAMp3V4DlaQN25l5_S8Zak3Mry2brOxhQ4";
            string range = "INPUTS!A1:D4";

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = spreadSheet.Rows[0].ColumnValue(2);
            thirtyDayColumn = spreadSheet.Rows[1].ColumnValue(2);

            //get index of thirty day column to use
            thirtyDay = NumberFromString(thirtyDayColumn) - NumberFromString(currentColumn);

            range = "GJ Capacity Forecast!" + currentColumn + "1:" + thirtyDayColumn + "13";
            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = " _(surplus of " + spreadSheet.Rows[8].ColumnValue(0) + ")_";
            if (float.Parse(spreadSheet.Rows[8].ColumnValue(0).ToString()) < 0)
            {
                currentColumn = " _(" + spreadSheet.Rows[8].ColumnValue(0).ToString().Replace("-","") + " people needed)_";
            }

            thirtyDayColumn = " _(surplus of " + spreadSheet.Rows[8].ColumnValue(thirtyDay) + ")_";
            if (float.Parse(spreadSheet.Rows[8].ColumnValue(thirtyDay).ToString()) < 0)
            {
                thirtyDayColumn = " _(" + spreadSheet.Rows[8].ColumnValue(thirtyDay).ToString().Replace("-","") + " people needed)_";
            }

            json = json + "*1800 GJ (based on " + spreadSheet.Rows[5].ColumnValue(0) + " rev)* " + spreadSheet.Rows[7].ColumnValue(0) + "/" + spreadSheet.Rows[6].ColumnValue(0) + currentColumn + " -- *30 day count:* " + spreadSheet.Rows[7].ColumnValue(thirtyDay) + "/" + spreadSheet.Rows[6].ColumnValue(thirtyDay) + thirtyDayColumn + "\n\n";



            //You Move Me
            //  get current and 30 day columns
            spreadSheetId = "1PH68RcBEgVDWjvCzf0wUym37kDZdGXvn_QsHqUNOCE8";
            range = "INPUTS!A1:D4";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = spreadSheet.Rows[0].ColumnValue(2);
            thirtyDayColumn = spreadSheet.Rows[1].ColumnValue(2);

            thirtyDay = NumberFromString(thirtyDayColumn) - NumberFromString(currentColumn);

            range = "YMM Capacity Forecast!" + currentColumn + "1:" + thirtyDayColumn + "13";
            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = " _(surplus of " + spreadSheet.Rows[8].ColumnValue(0) + ")_";
            if (float.Parse(spreadSheet.Rows[8].ColumnValue(0).ToString()) < 0)
            {
                currentColumn = " _(" + spreadSheet.Rows[8].ColumnValue(0).ToString().Replace("-", "") + " people needed)_";
            }

            thirtyDayColumn = " _(surplus of " + spreadSheet.Rows[8].ColumnValue(thirtyDay) + ")_";
            if (float.Parse(spreadSheet.Rows[8].ColumnValue(thirtyDay).ToString()) < 0)
            {
                thirtyDayColumn = " _(" + spreadSheet.Rows[8].ColumnValue(thirtyDay).ToString().Replace("-", "") + " people needed)_";
            }


            json = json + "*YMM (based on " + spreadSheet.Rows[5].ColumnValue(0) + " rev)* " + spreadSheet.Rows[7].ColumnValue(0) + "/" + spreadSheet.Rows[6].ColumnValue(0) + currentColumn + " -- *30 day count:* " + spreadSheet.Rows[7].ColumnValue(thirtyDay) + "/" + spreadSheet.Rows[6].ColumnValue(thirtyDay) + thirtyDayColumn + "\n\n";



            //Shack Shine
            //  get current and 30 day columns
            spreadSheetId = "1hVi7_DN2n0AsW6ZM75HzhdkcrqLAJkcZVCUzYPmkmvM";
            range = "INPUTS!A1:D4";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = spreadSheet.Rows[0].ColumnValue(2);
            thirtyDayColumn = spreadSheet.Rows[1].ColumnValue(2);

            thirtyDay = NumberFromString(thirtyDayColumn) - NumberFromString(currentColumn);

            range = "SS Capacity Forecast!" + currentColumn + "1:" + thirtyDayColumn + "13";
            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            currentColumn = " _(surplus of " + spreadSheet.Rows[7].ColumnValue(0) + ")_";
            if (float.Parse(spreadSheet.Rows[7].ColumnValue(0).ToString()) < 0)
            {
                currentColumn = " _(" + spreadSheet.Rows[7].ColumnValue(0).ToString().Replace("-", "") + " people needed)_";
            }

            thirtyDayColumn = " _(surplus of " + spreadSheet.Rows[7].ColumnValue(thirtyDay) + ")_";
            if (float.Parse(spreadSheet.Rows[7].ColumnValue(thirtyDay).ToString()) < 0)
            {
                thirtyDayColumn = " _(" + spreadSheet.Rows[7].ColumnValue(thirtyDay).ToString().Replace("-", "") + " people needed)_";
            }


            json = json + "*SS (based on " + spreadSheet.Rows[4].ColumnValue(0) + " rev)* " + spreadSheet.Rows[6].ColumnValue(0) + "/" + spreadSheet.Rows[5].ColumnValue(0) + currentColumn + " -- *30 day count:* " + spreadSheet.Rows[6].ColumnValue(thirtyDay) + "/" + spreadSheet.Rows[5].ColumnValue(thirtyDay) + thirtyDayColumn + "\n";



            json = json + "______________________________________________________\n";
            json = json + "\n   1-800-GOT-JUNK? Doc:   https://docs.google.com/spreadsheets/d/1IsdgX64l-JAMp3V4DlaQN25l5_S8Zak3Mry2brOxhQ4";
            json = json + "\n\n   You Move Me Reference Doc:   https://docs.google.com/spreadsheets/d/1PH68RcBEgVDWjvCzf0wUym37kDZdGXvn_QsHqUNOCE8";
            json = json + "\n\n   Shack Shine Reference Doc:   https://docs.google.com/spreadsheets/d/1hVi7_DN2n0AsW6ZM75HzhdkcrqLAJkcZVCUzYPmkmvM'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_hiringupdate").ToString().Replace("\\\\", "\\"));

        }


        public float ReturnSquareSwiped(string startDate, string endDate, string locationID, string locationAccessToken)
        {
            float totalJobs = 0;
            float swipedJobs = 0;

            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareSwiped> localList = new List<SquareSwiped>();
            SquareClient sc = new SquareClient();

            //localList = sc.RetrieveSquareSwiped(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));
            localList = sc.RetrieveSquareSwipedByDates(DateTime.Parse(startDate), DateTime.Parse(endDate), true, locationID, locationAccessToken);

            foreach (SquareSwiped x in localList)
            {
                totalJobs = totalJobs + 1;
                switch (x.ProcessingType.ToString().ToLower())
                {
                    case "keyed":
                        break;

                    default:
                        swipedJobs = swipedJobs + 1;
                        break;

                }
            }


            return swipedJobs / totalJobs;

        }



        public void UploadSlackUsers()
        {
            List<SlackUser> localList = new List<SlackUser>();
            //ConfigSettings cs = new ConfigSettings();
            SlackClient sc = new SlackClient();

            localList = sc.RetrieveUsers();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                //360 Slack Messages
                //string spreadSheetId = "1Xywsu3w8f6dQ_IVKykqG37twln3d0dnicm-ldyo34aU";

                //GJ Slack Messages
                string spreadSheetId = "1_HCikXcx-WzP7-l_PzCij6q3vPLhhWcZ5NiZcdDJeLQ";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UploadSlackUsers(spreadSheetId, localList);


            }
            catch (WebException ex)
            {

            }


        }



        public void SendToBuildingSecurity()
        {

            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            string scDailySchedule = cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_buildingsecurity").ToString().Replace("\\\\", "\\");

            //get schedule/employee information
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();
            List<Employee> jds = new List<Employee>();
            jds = gsc.ConnectSheetsBuildingSecurity();

            //send text message
            string scMessage = "{'text': 'Schedule for tomorrow: \n";
            foreach (Employee js in jds)
            {
                scMessage = scMessage + "   " + js.EmpName + " - " + js.EmpPhone + "\n";
            }
            scMessage = scMessage + "\n______________________________________________________\n";
            scMessage = scMessage + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1Lcd5o8sOq-lkxoPATiZHykwbPU7aNuyJLXVVUlbWPqc/edit#gid=0'}";

            //send to slack
            sc.SendMessage(scMessage, scDailySchedule);

        }


        //public void SendToSendMessages()
        //{
        //    RingCentralClient rc = new RingCentralClient();

        //    //get the numbers ....
        //    dbConnect dc = new dbConnect();
        //    dc.OpenMessageConnection();

        //    DataTable dt = new DataTable();

        //    SqlCommand cmd = new SqlCommand();
        //    cmd.Connection = dc.conn;
        //    cmd.CommandType = CommandType.StoredProcedure;
        //    cmd.CommandText = "spGetToSendMessages";

        //    SqlDataAdapter ds = new SqlDataAdapter(cmd);
        //    ds.Fill(dt);

        //    foreach (DataRow dr in dt.Rows)
        //    {

        //        //send the message ...
        //        rc.SendMMSRingCentral(dr[5].ToString(), dr[6].ToString(), dr[7].ToString(), dr[4].ToString(), "16122813895", dr[2].ToString(), dr[3].ToString());

        //        //if successful, delete the number
        //        cmd = new SqlCommand();
        //        cmd.Connection = dc.conn;
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.CommandText = "spDeleteToSend";
        //        cmd.Parameters.AddWithValue("@id", dr[0].ToString());

        //        cmd.ExecuteNonQuery();

        //    }

        //}

        public string RetrieveGoal(string segment)
        {
            string mtdRevenueGoal = "not found";
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();
            dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetDropDownListData";

            cmd.Parameters.Add("@segment", SqlDbType.NVarChar);
            cmd.Parameters["@segment"].Value = segment;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                mtdRevenueGoal = dt.Rows[0]["displayName"].ToString();
            }

            return mtdRevenueGoal;
        }

        public void RunNightlyChecklists(string bizUnitToUse, string slackChannel)
        {
            string dateToUse = DateTime.Now.ToShortDateString();
            float total = 0;
            float found = 0;
            SlackClient sc = new SlackClient();
            string output = "";

            dbConnect dc = new dbConnect();
            dc.OpenMessageConnection();

            DataTable dtResults = new DataTable();
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetMorningChecklistResults";

            cmd.Parameters.Add("@dateSent", SqlDbType.DateTime);
            cmd.Parameters["@dateSent"].Value = dateToUse;

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = bizUnitToUse;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                total = total + 1;
                dtResults = new DataTable();

                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetMorningChecklistY";

                cmd.Parameters.Add("@dateSent", SqlDbType.DateTime);
                cmd.Parameters["@dateSent"].Value = dateToUse;

                cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
                cmd.Parameters["@bizUnit"].Value = bizUnitToUse;

                cmd.Parameters.Add("@route", SqlDbType.NVarChar);
                cmd.Parameters["@route"].Value = dr["route"];

                ds = new SqlDataAdapter(cmd);
                ds.Fill(dtResults);

                if (dtResults.Rows.Count > 0)
                {
                    //results found
                    if (float.Parse(dr["countMorningChecklist"].ToString()) / float.Parse(dtResults.Rows[0]["countMorningChecklist"].ToString()) == 1)
                    {
                        found = found + 1;
                        output = output + "Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Morning Checklist Found\n";
                    }
                    else
                    {
                        //not found
                        output = output + "*Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Morning Checklist Not Found*\n";
                    }

                }
                else
                {
                    //no results found ... 0% completion rate
                    output = output + "*Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Morning Checklist Not Found*\n";
                }

            }


            string json = "{'text': '" + output + "\n\n" + dateToUse + " Morning Checklist Completion Rate: " + (found / total) * 100 + "%'}";
            sc.SendMessage(json, slackChannel);









            total = 0;
            found = 0;
            sc = new SlackClient();
            output = "";

            dtResults = new DataTable();
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetEveningChecklistResults";

            cmd.Parameters.Add("@dateSent", SqlDbType.DateTime);
            cmd.Parameters["@dateSent"].Value = dateToUse;

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = bizUnitToUse;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                total = total + 1;
                dtResults = new DataTable();

                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetEveningChecklistY";

                cmd.Parameters.Add("@dateSent", SqlDbType.DateTime);
                cmd.Parameters["@dateSent"].Value = dateToUse;

                cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
                cmd.Parameters["@bizUnit"].Value = bizUnitToUse;

                cmd.Parameters.Add("@route", SqlDbType.NVarChar);
                cmd.Parameters["@route"].Value = dr["route"];

                ds = new SqlDataAdapter(cmd);
                ds.Fill(dtResults);

                if (dtResults.Rows.Count > 0)
                {
                    //results found
                    if (float.Parse(dr["countEveningChecklist"].ToString()) / float.Parse(dtResults.Rows[0]["countEveningChecklist"].ToString()) == 1)
                    {
                        found = found + 1;
                        output = output + "Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Evening Checklist Found\n";
                    }
                    else
                    {
                        output = output + "*Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Evening Checklist Not Found*\n";
                    }

                }
                else
                {
                    //no results found ... 0% completion rate
                    output = output + "*Route " + dr["route"].ToString().Trim() + ": " + GetChecklistTeam(dateToUse, dr["route"].ToString(), bizUnitToUse) + " - Evening Checklist Not Found*\n";
                }

            }

            dc.CloseConnection();


            json = "{'text': '" + output + "\n\n" + dateToUse + " Evening Checklist Completion Rate: " + (found / total) * 100 + "%'}";
            sc.SendMessage(json, slackChannel);


        }


        public string GetChecklistTeam(string dateToUse, string route, string bizUnit)
        {
            string team = "";
            dbConnect dc = new dbConnect();
            dc.OpenMessageConnection();

            DataTable dtResults = new DataTable();
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetChecklistTeam";

            cmd.Parameters.Add("@dateSent", SqlDbType.DateTime);
            cmd.Parameters["@dateSent"].Value = dateToUse;

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = bizUnit;

            cmd.Parameters.Add("@route", SqlDbType.NVarChar);
            cmd.Parameters["@route"].Value = route;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                team = team + dr["routeTeam"] + ", ";

            }
            team = team.Substring(0, team.Length - 2);

            return team;

        }




        public AuditChecklist RetrieveJunkAuditChecklist(string auditID, string token, bool morning)
        {
            AuditChecklist localAD = new AuditChecklist();

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://api.safetyculture.io/audits/" + auditID);
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + token);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);

            string contents = _answer.ReadToEnd();

            //deserialize into a JSON object
            JObject jobj = JObject.Parse(contents);

            if (morning)
            {
                ////Pull the Audit Date label
                localAD.auditDate = "";
                try
                {
                    localAD.auditDate = ((string)jobj["header_items"][3]["responses"]["datetime"]).Substring(0, 10);
                }
                catch { }

                ////Pull the Audit Route label
                localAD.auditRoute = "";
                try
                {
                    localAD.auditRoute = (string)jobj["header_items"][7]["responses"]["selected"][0]["label"];
                }
                catch { }

                
            }
            else
            {
                ////Pull the Audit Date label
                localAD.auditDate = "";
                try
                {
                    localAD.auditDate = ((string)jobj["header_items"][2]["responses"]["datetime"]).Substring(0, 10);
                }
                catch { }


                ////Pull the Audit Route label
                localAD.auditRoute = "";
                try
                {
                    localAD.auditRoute = (string)jobj["header_items"][5]["responses"]["selected"][0]["label"];
                }
                catch { }
                
            }

            ////Pull the Audit ID label
            localAD.auditID = auditID;

            return localAD;
        }


        public AuditChecklist RetrieveAuditChecklist(string auditID, string token)
        {
            AuditChecklist localAD = new AuditChecklist();

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://api.safetyculture.io/audits/" + auditID);
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + token);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);

            string contents = _answer.ReadToEnd();

            //deserialize into a JSON object
            JObject jobj = JObject.Parse(contents);

            ////Pull the Audit Date label
            localAD.auditDate = "";
            try
            {
                localAD.auditDate = ((string)jobj["header_items"][2]["responses"]["datetime"]).Substring(0, 10);
            }
            catch { }

            ////Pull the Audit Route label
            localAD.auditRoute = "";
            try
            {
                localAD.auditRoute = (string)jobj["header_items"][4]["responses"]["selected"][0]["label"];
            }
            catch { }
            

            ////Pull the Audit ID label
            localAD.auditID = auditID;

            return localAD;
        }


        public AuditData RetrieveiAuditorAuditCustomer(string auditID, string token)
        {
            AuditData localAD = new AuditData();

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://api.safetyculture.io/audits/" + auditID);
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + token);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);

            string contents = _answer.ReadToEnd();

            //deserialize into a JSON object
            JObject jobj = JObject.Parse(contents);

            //Pull the Work Order ID label
            localAD.workOrderID = "";
            try
            {
                localAD.workOrderID = (string)jobj["header_items"][2]["responses"]["text"];
            }
            catch { }

            //Pull the Client Name label
            localAD.customerName = "";
            try
            {
                localAD.customerName = (string)jobj["header_items"][3]["responses"]["text"];
            }
            catch { }

            //Pull the Damage Happened Load label
            localAD.damageHappenedLoad = "N/A";
            try
            {
                localAD.damageHappenedLoad = (string)jobj["items"][36]["responses"]["text"];
            }
            catch { }

            //Pull the Damage Prevention Load label
            localAD.damagePreventionLoad = "N/A";
            try
            {
                localAD.damagePreventionLoad = (string)jobj["items"][38]["responses"]["text"];
            }
            catch { }

            //Pull the Damage Happened Unload label
            localAD.damageHappenedUnload = "N/A";
            try
            {
                localAD.damageHappenedUnload = (string)jobj["items"][84]["responses"]["text"];
            }
            catch { }

            //Pull the Damage Prevention Unload label
            localAD.damagePreventionUnload = "N/A";
            try
            {
                localAD.damagePreventionUnload = (string)jobj["items"][86]["responses"]["text"];
            }
            catch { }

            return localAD;
        }
        

        public List<Audit> RetrieveiAuditorAudits(string token, string templateID)
        {

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://api.safetyculture.io/audits/search?order=desc&completed=true&limit=20");
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + token);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);


            AuditsRoot deserializedAudit = new AuditsRoot();
            MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(_answer.ReadToEnd()));
            DataContractJsonSerializer ser = new DataContractJsonSerializer(deserializedAudit.GetType());
            deserializedAudit = ser.ReadObject(ms) as AuditsRoot;
            ms.Close();

            Audit r = new Audit();
            List<Audit> localList = new List<Audit>();
            try
            {
                for (int i = 0; i < deserializedAudit.count; i++)
                {
                    if (deserializedAudit.audits[i].template_id.ToString().ToLower() == templateID.ToLower())
                    {
                        r = new Audit();
                        r.audit_id = deserializedAudit.audits[i].audit_id;
                        r.template_id = deserializedAudit.audits[i].template_id;
                        localList.Add(r);

                    }
                }
            }
            catch
            {
            }

            return localList;
        }



        public float RetrieveWarehouseSale()
        {
            float resale = 0;
            DateTime dateValue;
            int currentMonth = 1;
            int currentYear = 1;

            ShackWagesMessageModel mwm = new ShackWagesMessageModel();

            DateTime now = DateTime.Now;
            currentYear = now.Year;
            switch (now.Day)
            {
                case 1:
                    currentMonth = now.AddDays(-1).Month;
                    break;
                default:
                    currentMonth = now.Month;
                    break;
            }

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1kqxIC2S8iJRDlOUECNjtpCMEQqu8_CVXF8mv77sw7OA";
                string range = "Full Sheet!A1:H5000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (DateTime.TryParse(row.ColumnValue(6).ToString().Trim(), out dateValue))
                        {
                            if ((dateValue.Month == currentMonth) && (dateValue.Year == currentYear))
                            {
                                if (!string.IsNullOrEmpty(row.ColumnValue(7).ToString().Trim().Replace(" ", "").Replace("$", "").Replace("-", "")))
                                {
                                    resale = resale + float.Parse(row.ColumnValue(7).ToString().Trim().Replace(" ", "").Replace("$", ""));
                                }
                            }
                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }

            }
            catch (WebException ex)
            {

            }

            return resale;

        }
        
        public double CalculateRevenuePace(string startDate, string endDate)
        {
            double pace = 0;
            int daysInMonth = 0;
            double avgPerDay = 0;

            DateTime tempDate;

            double sunTotal = 0;
            double sunCount = 0;
            double sunAvg = 0;
            double monTotal = 0;
            double monCount = 0;
            double monAvg = 0;
            double tueTotal = 0;
            double tueCount = 0;
            double tueAvg = 0;
            double wedTotal = 0;
            double wedCount = 0;
            double wedAvg = 0;
            double thuTotal = 0;
            double thuCount = 0;
            double thuAvg = 0;
            double friTotal = 0;
            double friCount = 0;
            double friAvg = 0;
            double satTotal = 0;
            double satCount = 0;
            double satAvg = 0;


            //get an average / day thus far for the month
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalRev"].ToString()))
                {
                    avgPerDay = Double.Parse(dt.Rows[0][0].ToString()) / DateTime.Parse(endDate).Day;
                }
            }




            //get average per day of the week ... if average = 0, then set to the average/day thus far
            dc = new dbConnect();
            dc.OpenConnection();

            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetRevGroupedByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr in dt.Rows)
                {
                    tempDate = DateTime.Parse(dr[0].ToString());

                    switch (tempDate.DayOfWeek)
                    {
                        case DayOfWeek.Sunday:
                            sunTotal = sunTotal + Double.Parse(dr[1].ToString());
                            sunCount = sunCount + 1;
                            break;
                        case DayOfWeek.Monday:
                            monTotal = monTotal + Double.Parse(dr[1].ToString());
                            monCount = monCount + 1;
                            break;
                        case DayOfWeek.Tuesday:
                            tueTotal = tueTotal + Double.Parse(dr[1].ToString());
                            tueCount = tueCount + 1;
                            break;
                        case DayOfWeek.Wednesday:
                            wedTotal = wedTotal + Double.Parse(dr[1].ToString());
                            wedCount = wedCount + 1;
                            break;
                        case DayOfWeek.Thursday:
                            thuTotal = thuTotal + Double.Parse(dr[1].ToString());
                            thuCount = thuCount + 1;
                            break;
                        case DayOfWeek.Friday:
                            friTotal = friTotal + Double.Parse(dr[1].ToString());
                            friCount = friCount + 1;
                            break;
                        case DayOfWeek.Saturday:
                            satTotal = satTotal + Double.Parse(dr[1].ToString());
                            satCount = satCount + 1;
                            break;

                    }

                }

                if (sunCount > 0) {sunAvg = sunTotal / sunCount;}
                if (monCount > 0) { monAvg = monTotal / monCount; }
                if (tueCount > 0) { tueAvg = tueTotal / tueCount; }
                if (wedCount > 0) { wedAvg = wedTotal / wedCount; }
                if (thuCount > 0) { thuAvg = thuTotal / thuCount; }
                if (friCount > 0) { friAvg = friTotal / friCount; }
                if (satCount > 0) { satAvg = satTotal / satCount; }


                if (sunAvg == 0) { sunAvg = avgPerDay; }
                if (monAvg == 0) { monAvg = avgPerDay; }
                if (tueAvg == 0) { tueAvg = avgPerDay; }
                if (wedAvg == 0) { wedAvg = avgPerDay; }
                if (thuAvg == 0) { thuAvg = avgPerDay; }
                if (friAvg == 0) { friAvg = avgPerDay; }
                if (satAvg == 0) { satAvg = avgPerDay; }

            }



            //get days in current month
            switch (DateTime.Parse(endDate).Month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    daysInMonth = 31;
                    break;
                case 2:
                    daysInMonth = 28;
                    break;
                case 4:
                case 6:
                case 9:
                case 11:
                    daysInMonth = 30;
                    break;
            }


            //cycle through rest of month and add average to total
            for (int i = 1; i <= daysInMonth - DateTime.Parse(endDate).Day; i++)
            {
                tempDate = DateTime.Parse(endDate).AddDays(i);

                switch (tempDate.DayOfWeek)
                {
                    case DayOfWeek.Sunday:
                        sunTotal = sunTotal + sunAvg;
                        sunCount = sunCount + 1;
                        break;
                    case DayOfWeek.Monday:
                        monTotal = monTotal + monAvg;
                        monCount = monCount + 1;
                        break;
                    case DayOfWeek.Tuesday:
                        tueTotal = tueTotal + tueAvg;
                        tueCount = tueCount + 1;
                        break;
                    case DayOfWeek.Wednesday:
                        wedTotal = wedTotal + wedAvg;
                        wedCount = wedCount + 1;
                        break;
                    case DayOfWeek.Thursday:
                        thuTotal = thuTotal + thuAvg;
                        thuCount = thuCount + 1;
                        break;
                    case DayOfWeek.Friday:
                        friTotal = friTotal + friAvg;
                        friCount = friCount + 1;
                        break;
                    case DayOfWeek.Saturday:
                        satTotal = satTotal + satAvg;
                        satCount = satCount + 1;
                        break;

                }


            }


            pace = sunTotal + monTotal + tueTotal + wedTotal + thuTotal + friTotal + satTotal;

            return pace;

        }







    }
}
