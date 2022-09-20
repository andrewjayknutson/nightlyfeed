using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utilities360Wow;
using System.Configuration;
using NightlyRouteToSlack.iAuditorAudits;
using System.Data;
using System.Data.SqlClient;
using Google.Apis.Sheets.v4;
using System.Net;
using System.IO;
using System.Text;


namespace NightlyRouteToSlack.Utilities
{
    public class JunkUtilities
    {

        public void RunGJDailyHoursWorkedDump()
        {
            //this will run every day at 2am so 
            //  startDate and endDate = the day prior
            string startDate = DateTime.Now.AddDays(-1).ToShortDateString();
            string endDate = DateTime.Now.AddDays(-1).ToShortDateString();
            string weekStartDate = "";

            //to do:
            //  - get revenue for the four week period starting in cell B1
            //  - import that revenue number to cell K1 on tab reporting


            //  - import hours from the previous day to tab raw_data
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetWagesV2";

            cmd.Parameters.Add("@startDate", SqlDbType.DateTime);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1VEtf2Smn5VhxyfCBnad5l-G-nltVM5QBnc8Zd4QRwqc";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                //reader.UpdateJunkHoursWorked(spreadSheetId, dt);




                //  - get the starting date from cell B1
                string range = "reporting!A1:B2";

                reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(spreadSheetId, range);

                //overall NPS score
                weekStartDate = spreadSheet.Rows[1].ColumnValue(1).ToString();


            }
            catch (WebException ex)
            {

            }











        }


        public void RunGJNPSScorecard()
        {

            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            int x = 0;
            string json = "{'text': '*DAILY NPS TRACKING*\n\n";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages m = new MoveDailyWages();

            m = gc.ConnectSheetsJunkNPSScorecard("105183BgGKXcdYSg1HseDXl7sOlNhAqauDr8fJDaXxNw");

            json = json + "*" + m.startDate + " - " + m.endDate + "*\n";
            //json = json + "   Thru Date: " + m.thruDate + "\n";
            //json = json + "   Last Updated: " + m.lastUpdated + "\n";
            //json = json + "   Overall NPS: " + m.score + "         *Goal: " + m.goal + "*\n\n";
            json = json + "   *Team Members*\n\n";


            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                json = json + "   " + mdw.StartDate.Replace("'", "") + ":   " + mdw.EndDate.Replace("'", "") + "    " + mdw.CrewLead + " response(s)\n";
            }
            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/105183BgGKXcdYSg1HseDXl7sOlNhAqauDr8fJDaXxNw/edit#gid=723168239'}";

            //NPS Scorecard Slack Channel
            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junknpsscorecard").ToString().Replace("\\\\", "\\"));




        }


        public void RunSalesLeadAJS()
        {
            string json = "";
            bool firstOne = true;
            string startDate = DateTime.Now.AddDays(-7).ToShortDateString();
            string endDate = DateTime.Now.AddDays(-1).ToShortDateString();

            //startDate = "5/1/2020";
            //endDate = "5/31/2020";

            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            json = "{'text': '*Sales Lead Residential AJS*\n  *" + startDate + " - " + endDate + "*\n\n  *ABOVE $400 GOAL*\n";

            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spReport_SalesLeadAJS";

            cmd.Parameters.Add("@startDate", SqlDbType.DateTime);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            foreach (DataRow dr in dt.Rows)
            {
                if (float.Parse(dr["AJS"].ToString()) < 400)
                {
                    if (firstOne)
                    {
                        firstOne = false;
                        json = json + "\n\n  *BELOW $400 GOAL*\n";
                    }
                }

                json = json + "  " + dr["firstName"].ToString() + " " + dr["lastName"].ToString() + " " + float.Parse(dr["AJS"].ToString()).ToString("C") + "   (" + float.Parse(dr["Totalnet"].ToString()).ToString("C") + " total revenue / " + float.Parse(dr["CountNet"].ToString()).ToString() + " jobs)" + "\n";
            }



            //json = json + "______________________________________________________\n";
            //json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ'}";
            json = json + "\n'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gjsalesleadajs").ToString().Replace("\\\\", "\\"));

        }

        public void RunDailyHealthCheckImport()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            //get schedule/employee information
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();
            List<JunkDailySchedule> jds = new List<JunkDailySchedule>();


            List<JunkDailySchedule> jdsWest = new List<JunkDailySchedule>();
            jdsWest = gsc.ConnectSheetsJunkDailySchedule("West Schedule!", "L1:N100");
            foreach (JunkDailySchedule js in jdsWest)
            {
                JunkDailySchedule jsTemp = new JunkDailySchedule();
                jsTemp.EmpName = js.EmpName;
                jsTemp.EmpPhone = js.EmpPhone;
                jsTemp.EmpStartTime = js.EmpStartTime;
                jsTemp.EmpWorkType = js.EmpWorkType;

                jds.Add(jsTemp);
            }


            List<JunkDailySchedule> jdsEast = new List<JunkDailySchedule>();
            jdsEast = gsc.ConnectSheetsJunkDailySchedule("East Schedule!", "L1:N100");
            foreach (JunkDailySchedule js in jdsEast)
            {
                JunkDailySchedule jsTemp = new JunkDailySchedule();
                jsTemp.EmpName = js.EmpName;
                jsTemp.EmpPhone = js.EmpPhone;
                jsTemp.EmpStartTime = js.EmpStartTime;
                jsTemp.EmpWorkType = js.EmpWorkType;

                jds.Add(jsTemp);
            }


            //import to daily sheet
            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            reader.ImportDailyHealthCheckSheet("1dFOiu7ZUEs2EpgQxEFmTjO599AFQsw5j5vYqa9rn7Uc", jds, "1800GJ");



        }


        public void JunkOTAwarenessUpdate()
        {
            string startDate = "";
            string endDate = "";
            string empList = "";
            string json = "";
            int count = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);


            //get dates to run
            string spreadSheetId = "1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ";
            string range = "OT Awareness!A3:A5";

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            startDate = spreadSheet.Rows[0].ColumnValue(0);
            endDate = spreadSheet.Rows[1].ColumnValue(0);
            //startDate = "1/12/2020";
            //endDate = "1/18/2020";


            //get people working today
            range = "Sheet1!A1:D63";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //Read through and populate a list of employees
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                try
                {
                    if (row.ColumnValue(1).ToString().Trim() != "") 
                    {
                        if (row.ColumnValue(1).ToString().Trim().ToLower() != "name" & row.ColumnValue(1).ToString().Trim().Substring(0, 4).ToLower() != "truc")
                        {
                            empList = empList + "," + row.ColumnValue(1).ToString().Trim();
                        }
                    }
                }
                catch (WebException ex)
                {

                }
            }




            //Go out and get wage data
            UniformData ud = new UniformData();
            List<UniformData> sd = new List<UniformData>();
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetOTAwareness";

            cmd.Parameters.Add("@startDate", SqlDbType.DateTime);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            foreach (DataRow dr in dt.Rows)
            {
                ud = new UniformData();

                ud.ItemName = dr["firstName"].ToString();
                ud.ItemReorder = dr["lastName"].ToString();
                ud.ItemReorderAmount = dr["totalHours"].ToString();

                sd.Add(ud);

            }


            //update 
            reader = new GoogleSpreadSheetReader(googleService);
            reader.UpdateJunkOTAwareness(sd);

            json = "{'text': '*Employees Close to Overtime*\n";

            //check to see if any are within 10 hours of 40 and working today ... if so, send to slack
            foreach (UniformData localUD in sd)
            {
                if (float.Parse(localUD.ItemReorderAmount) >= 30)
                {
                    //check if on the schedule for today
                    if (empList.IndexOf(localUD.ItemName + " " + localUD.ItemReorder) > -1)
                    {
                        //found ... send to slack
                        json = json + "      " + localUD.ItemName + " " + localUD.ItemReorder + "  " + localUD.ItemReorderAmount + "   *Scheduled Today*\n";
                        count = count + 1;
                    }
                    else 
                    {
                        json = json + "      " + localUD.ItemName + " " + localUD.ItemReorder + "  " + localUD.ItemReorderAmount + "\n";

                    }


                }

            }

            json = json + "______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ'}";

            if (count > 0)
            {
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junkweeklyscorecard").ToString().Replace("\\\\", "\\"));
            }











            //*****************
            //You Move Me
            //*****************
            ud = new UniformData();
            List<UniformData> empHoursList = new List<UniformData>();
            empList = "";

            //get dates to run
            spreadSheetId = "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I";
            range = "OT Awareness!A1:C50";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //Read through and populate a list of employees
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                if (row.ColumnValue(1).ToString().Trim() != "" & row.ColumnValue(2).ToString().Trim().ToLower() != "hours")
                {
                    ud = new UniformData();
                    ud.ItemName = row.ColumnValue(1).ToString().Trim().Replace(" - ", "").Replace("Sr Mover", "").Replace("Mover", "").Replace("Sr Crew Lead", "").Replace("Crew Lead", "");
                    ud.ItemReorder = row.ColumnValue(2).ToString();

                    empHoursList.Add(ud);
                }
            }



            //Get Today Workers
            spreadSheetId = "1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs";
            range = "Daily Schedule!A1:D70";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //Read through and populate a list of employees
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                if (row.ColumnValue(1).ToString().Trim() != "")
                {
                    if (row.ColumnValue(1).ToString().Trim().ToLower() != "name" & row.ColumnValue(1).ToString().Trim().Substring(0, 4).ToLower() != "crew")
                    {
                        empList = empList + "," + row.ColumnValue(1).ToString().Trim();
                    }
                }
            }


            count = 0;
            json = "{'text': '*Employees Close to Overtime*\n";

            //check to see if any are within 10 hours of 40 and working today ... if so, send to slack
            foreach (UniformData localEmp in empHoursList)
            {
                if (float.Parse(localEmp.ItemReorder) >= 30)
                {
                    //check if on the schedule for today
                    if (empList.IndexOf(localEmp.ItemName) > -1)
                    {
                        //found ... send to slack
                        json = json + "      " + localEmp.ItemName + "  " + localEmp.ItemReorder + "   *Scheduled Today*\n";
                    }
                    else
                    {
                        json = json + "      " + localEmp.ItemName + " " + localEmp.ItemReorder + "\n";

                    }
                    count = count + 1;

                }

            }

            json = json + "______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs'}";

            if (count > 0)
            {
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_moveweeklyscorecard").ToString().Replace("\\\\", "\\"));
            }











            //*****************
            //Shack Shine
            //*****************
            ud = new UniformData();
            empHoursList = new List<UniformData>();
            empList = "";

            //get dates to run
            spreadSheetId = "1ttCMCGo-C5ozvlgWgMvQKkSY0Ujw4hO_yJlL3qLOn_Q";
            range = "OT Awareness!A1:C50";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //Read through and populate a list of employees
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                if (row.ColumnValue(1).ToString().Trim() != "" & row.ColumnValue(2).ToString().Trim().ToLower() != "hours")
                {
                    ud = new UniformData();
                    ud.ItemName = row.ColumnValue(1).ToString().Trim().Replace(" - ", "").Replace("Lead Tech", "").Replace("Tech", "");
                    ud.ItemReorder = row.ColumnValue(2).ToString();

                    empHoursList.Add(ud);
                }
            }



            ////Get Today Workers
            //spreadSheetId = "1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs";
            //range = "Daily Schedule!A1:D70";

            //reader = new GoogleSpreadSheetReader(googleService);
            //spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            ////Read through and populate a list of employees
            //foreach (SpreadSheetRow row in spreadSheet.Rows)
            //{
            //    if (row.ColumnValue(1).ToString().Trim() != "")
            //    {
            //        if (row.ColumnValue(1).ToString().Trim().ToLower() != "name" & row.ColumnValue(1).ToString().Trim().Substring(0, 4).ToLower() != "crew")
            //        {
            //            empList = empList + "," + row.ColumnValue(1).ToString().Trim();
            //        }
            //    }
            //}


            count = 0;
            json = "{'text': '*Employees Close to Overtime*\n";

            //check to see if any are within 10 hours of 40 and working today ... if so, send to slack
            foreach (UniformData localEmp in empHoursList)
            {
                if (float.Parse(localEmp.ItemReorder) >= 30)
                {
                    ////check if on the schedule for today
                    //if (empList.IndexOf(localEmp.ItemName) > -1)
                    //{
                    //    //found ... send to slack
                    //    json = json + "      " + localEmp.ItemName + "  " + localEmp.ItemReorder + "   *Scheduled Today*\n";
                    //}
                    //else
                    //{
                        json = json + "      " + localEmp.ItemName + " " + localEmp.ItemReorder + "\n";

                    //}
                    count = count + 1;
                }

            }

            json = json + "______________________________________________________\n";
            //json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs'}";
            json = json + "\n'}";

            if (count > 0)
            {
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackweeklyscorecard").ToString().Replace("\\\\", "\\"));
            }









        }


        public void JunkDailyRecon()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string json = "";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWage mdw = new MoveDailyWage();

            mdw = gc.ConnectSheetsGJDailyRecon("114T2kCA174JbPTuW_kd_x1LF-N7eEAVhpBkj6NHmGbE");

            json = "{'text': '*Daily Recon Results to Pipeline*\n";
            json = json + "Route Date: " + mdw.WagePercentage + "\n";
            if (mdw.CrewLead.ToString().Trim() == "")
            {
                json = json + "Closer: *MISSING*\n";
            }
            else 
            {
                json = json + "Closer: " + mdw.CrewLead + "\n";
            }

            if (mdw.RouteNumber.ToString().Trim() == "")
            {
                json = json + "Pipeline: *MISSING*\n";
            }
            else 
            {
                json = json + "Pipeline: " + mdw.RouteNumber + "\n";
            }

            if (mdw.RouteDate.ToString().Trim() == "")
            {
                json = json + "RM: *MISSING*\n";
            }
            else
            {
                json = json + "RM: " + mdw.RouteDate + "       " + mdw.RPH + "\n";
            }


            json = json + "______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/114T2kCA174JbPTuW_kd_x1LF-N7eEAVhpBkj6NHmGbE/edit?pli=1#gid=511603474'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junkdailyrecon").ToString().Replace("\\\\", "\\"));

        }



        public void JunkScorecardDaily()
        {
            CommonClient cc = new CommonClient();
            UtilityClass uc = new UtilityClass();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            float yesterdayRevenue = 0;
            float yesterdayJobs = 0;
            float yesterdayAJS = 0;
            float yesterdayRPH = 0;
            float yesterdayDirectWage = 0;
            float yesterdayDirectWagePercentage = 0;

            float yesterdayBonusWage = 0;
            float yesterdayBonusWagePercentage = 0;
            float yesterdayTrainerWage = 0;
            float yesterdayTrainerWagePercentage = 0;
            float yesterdayTruckWage = 0;
            float yesterdayTruckWagePercentage = 0;
            float yesterdayClericalWage = 0;
            float yesterdayClericalWagePercentage = 0;
            float yesterdayTrainingWage = 0;
            float yesterdayTrainingWagePercentage = 0;
            float yesterdayMeetingWage = 0;
            float yesterdayMeetingWagePercentage = 0;
            float yesterdaySupportWage = 0;
            float yesterdaySupportWagePercentage = 0;
            float yesterdayPointWage = 0;
            float yesterdayPointWagePercentage = 0;
            float yesterdayFleetWage = 0;
            float yesterdayFleetWagePercentage = 0;
            float yesterdayWarehouseWage = 0;
            float yesterdayWarehouseWagePercentage = 0;

            float yesterdayIndirectWage = 0;
            float yesterdayIndirectWagePercentage = 0;
            float yesterdaySwiped = 0;
            float mtdRevenue = 0;
            float lastyearMTDRevenue = 0;
            float mtdRevenuePace = 0;
            float mtdRevenueGoal = 0;
            float mtdJobs = 0;
            float mtdAJS = 0;
            float mtdRPH = 0;
            float mtdDirectWage = 0;
            float mtdDirectWagePercentage = 0;
            float mtdIndirectWage = 0;
            float mtdIndirectWagePercentage = 0;
            float mtdSwiped = 0;
            float mtdOpenAR = 0;

            string json = "{'text': '";

            string yesterday = DateTime.Today.AddDays(-1).ToShortDateString();
            string firstOfMonthDate = cc.GetStartDate();
            string endOfMonthDate = cc.GetEOMDate();


            yesterday = "10/21/2021";
            firstOfMonthDate = "10/1/2021";
            endOfMonthDate = "10/31/2021";


            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            //**** Yesterday Revenue
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = yesterday;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = yesterday;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalRev"].ToString()))
                {
                    yesterdayRevenue = float.Parse(dt.Rows[0]["totalRev"].ToString());
                    yesterdayJobs = float.Parse(dt.Rows[0]["countRev"].ToString());
                }

                yesterdayAJS = yesterdayRevenue / yesterdayJobs;
            }


            //**** Yesterday Direct Wages ****
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetWagesByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = yesterday;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = yesterday;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["totalPay"].ToString() == "") { }
                else
                {
                    yesterdayDirectWage = float.Parse(dt.Rows[0]["totalDirectLabor"].ToString());
                    yesterdayDirectWagePercentage = yesterdayDirectWage / yesterdayRevenue;
                    yesterdayIndirectWage = float.Parse(dt.Rows[0]["totalIndirectLabor"].ToString());
                    yesterdayIndirectWagePercentage = yesterdayIndirectWage / yesterdayRevenue;
                    yesterdayRPH = yesterdayRevenue / float.Parse(dt.Rows[0]["totalDirectHours"].ToString());
                }
            }


            //**** Yesterday Direct Wage Types ****
            //dt = new DataTable();

            //cmd = new SqlCommand();
            //cmd.Connection = dc.conn;
            //cmd.CommandType = CommandType.StoredProcedure;
            //cmd.CommandText = "spAutomated_GetWagesTypeByDates";

            //cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            //cmd.Parameters["@startDate"].Value = yesterday;

            //cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            //cmd.Parameters["@endDate"].Value = yesterday;

            //ds = new SqlDataAdapter(cmd);
            //ds.Fill(dt);
            //dc.CloseConnection();

            //if (dt.Rows.Count > 0)
            //{
            //    if (dt.Rows[0]["totalTruckWages"].ToString() == "") { }
            //    else
            //    {
            //        yesterdayTrainerWage = float.Parse(dt.Rows[0]["totalTrainerWages"].ToString());
            //        yesterdayTrainerWagePercentage = yesterdayTrainerWage / yesterdayRevenue;
            //        yesterdayBonusWage = float.Parse(dt.Rows[0]["totalBonusWages"].ToString());
            //        yesterdayBonusWagePercentage = yesterdayBonusWage / yesterdayRevenue;
            //        yesterdayTruckWage = float.Parse(dt.Rows[0]["totalTruckWages"].ToString());
            //        yesterdayTruckWagePercentage = yesterdayTruckWage / yesterdayRevenue;
            //        yesterdayClericalWage = float.Parse(dt.Rows[0]["totalClericalWages"].ToString());
            //        yesterdayClericalWagePercentage = yesterdayClericalWage / yesterdayRevenue;
            //        yesterdayTrainingWage = float.Parse(dt.Rows[0]["totalTrainingWages"].ToString());
            //        yesterdayTrainingWagePercentage = yesterdayTrainingWage / yesterdayRevenue;
            //        yesterdayMeetingWage = float.Parse(dt.Rows[0]["totalMeetingWages"].ToString());
            //        yesterdayMeetingWagePercentage = yesterdayMeetingWage / yesterdayRevenue;
            //        yesterdaySupportWage = float.Parse(dt.Rows[0]["totalSupportWages"].ToString());
            //        yesterdaySupportWagePercentage = yesterdaySupportWage / yesterdayRevenue;
            //        yesterdayPointWage = float.Parse(dt.Rows[0]["totalPointWages"].ToString());
            //        yesterdayPointWagePercentage = yesterdayPointWage / yesterdayRevenue;
            //        yesterdayFleetWage = float.Parse(dt.Rows[0]["totalFleetWages"].ToString());
            //        yesterdayFleetWagePercentage = yesterdayFleetWage / yesterdayRevenue;
            //        yesterdayWarehouseWage = float.Parse(dt.Rows[0]["totalWarehouseWages"].ToString());
            //        yesterdayWarehouseWagePercentage = yesterdayWarehouseWage / yesterdayRevenue;
            //    }
            //}



            //yesterdaySwiped = uc.ReturnSquareSwiped(yesterday, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));


            //need to get cells B10 & B13 from https://docs.google.com/spreadsheets/d/1ipm46yj2yD1s_28VL80gm4IpLu1s2XsV60tGSL1khKA/edit#gid=696941223



            json = json + "*DAILY SCORECARD*\n\n*" + yesterday + "*\n";
            json = json + "  Revenue: " + yesterdayRevenue.ToString("C") + "\n";
            json = json + "  AJS: " + yesterdayAJS.ToString("C") + "\n";
            json = json + "  RPH: " + yesterdayRPH.ToString("C") + "\n";
            json = json + "  Direct Wage: " + yesterdayDirectWagePercentage.ToString("P") + "\n";
            json = json + "  SG&A: " + yesterdayDirectWagePercentage.ToString("P") + "\n";
            //json = json + "    Truck: " + yesterdayTruckWagePercentage.ToString("P") + "\n";
            //json = json + "    Clerical: " + yesterdayClericalWagePercentage.ToString("P") + "\n";
            //json = json + "    Meeting: " + yesterdayMeetingWagePercentage.ToString("P") + "\n";
            //json = json + "    Training: " + yesterdayTrainingWagePercentage.ToString("P") + "\n";
            //json = json + "    Support: " + yesterdaySupportWagePercentage.ToString("P") + "\n";
            //json = json + "    Warehouse: " + yesterdayWarehouseWagePercentage.ToString("P") + "\n";
            //json = json + "    Daily Bonus: " + yesterdayBonusWagePercentage.ToString("P") + "\n";
            //json = json + "    Trainer: " + yesterdayTrainerWagePercentage.ToString("P") + "\n";
            //json = json + "  Indirect Wage: " + yesterdayIndirectWagePercentage.ToString("P") + "\n";
            //json = json + "    Point: " + yesterdayPointWagePercentage.ToString("P") + "\n";
            //json = json + "    Fleet: " + yesterdayFleetWagePercentage.ToString("P") + "\n";
            //json = json + "  Swiped: " + yesterdaySwiped.ToString("P") + "\n";




            //**** MTD Revenue
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = firstOfMonthDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = yesterday;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalRev"].ToString()))
                {
                    mtdRevenue = float.Parse(dt.Rows[0]["totalRev"].ToString());
                    mtdJobs = float.Parse(dt.Rows[0]["countRev"].ToString());
                }

                mtdAJS = mtdRevenue / mtdJobs;
            }


            //**** MTD Direct Wages ****
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetWagesByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = firstOfMonthDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = yesterday;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["totalPay"].ToString() == "") { }
                else
                {
                    mtdDirectWage = float.Parse(dt.Rows[0]["totalDirectLabor"].ToString());
                    mtdDirectWagePercentage = mtdDirectWage / mtdRevenue;
                    mtdIndirectWage = float.Parse(dt.Rows[0]["totalIndirectLabor"].ToString());
                    mtdIndirectWagePercentage = mtdIndirectWage / mtdRevenue;
                    mtdRPH = mtdRevenue / float.Parse(dt.Rows[0]["totalDirectHours"].ToString());
                }
            }


            //**** Yesterday Direct Wage Types ****
            //dt = new DataTable();

            //cmd = new SqlCommand();
            //cmd.Connection = dc.conn;
            //cmd.CommandType = CommandType.StoredProcedure;
            //cmd.CommandText = "spAutomated_GetWagesTypeByDates";

            //cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            //cmd.Parameters["@startDate"].Value = firstOfMonthDate;

            //cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            //cmd.Parameters["@endDate"].Value = yesterday;

            //ds = new SqlDataAdapter(cmd);
            //ds.Fill(dt);
            //dc.CloseConnection();

            //if (dt.Rows.Count > 0)
            //{
            //    if (dt.Rows[0]["totalTruckWages"].ToString() == "") { }
            //    else
            //    {
            //        yesterdayTrainerWage = float.Parse(dt.Rows[0]["totalTrainerWages"].ToString());
            //        yesterdayTrainerWagePercentage = yesterdayTrainerWage / mtdRevenue;
            //        yesterdayBonusWage = float.Parse(dt.Rows[0]["totalBonusWages"].ToString());
            //        yesterdayBonusWagePercentage = yesterdayBonusWage / mtdRevenue;
            //        yesterdayTruckWage = float.Parse(dt.Rows[0]["totalTruckWages"].ToString());
            //        yesterdayTruckWagePercentage = yesterdayTruckWage / mtdRevenue;
            //        yesterdayClericalWage = float.Parse(dt.Rows[0]["totalClericalWages"].ToString());
            //        yesterdayClericalWagePercentage = yesterdayClericalWage / mtdRevenue;
            //        yesterdayTrainingWage = float.Parse(dt.Rows[0]["totalTrainingWages"].ToString());
            //        yesterdayTrainingWagePercentage = yesterdayTrainingWage / mtdRevenue;
            //        yesterdayMeetingWage = float.Parse(dt.Rows[0]["totalMeetingWages"].ToString());
            //        yesterdayMeetingWagePercentage = yesterdayMeetingWage / mtdRevenue;
            //        yesterdaySupportWage = float.Parse(dt.Rows[0]["totalSupportWages"].ToString());
            //        yesterdaySupportWagePercentage = yesterdaySupportWage / mtdRevenue;
            //        yesterdayPointWage = float.Parse(dt.Rows[0]["totalPointWages"].ToString());
            //        yesterdayPointWagePercentage = yesterdayPointWage / mtdRevenue;
            //        yesterdayFleetWage = float.Parse(dt.Rows[0]["totalFleetWages"].ToString());
            //        yesterdayFleetWagePercentage = yesterdayFleetWage / mtdRevenue;
            //        yesterdayWarehouseWage = float.Parse(dt.Rows[0]["totalWarehouseWages"].ToString());
            //        yesterdayWarehouseWagePercentage = yesterdayWarehouseWage / mtdRevenue;
            //    }
            //}





            //**** Monthly Goal Revenue
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetDropDownListData";

            cmd.Parameters.Add("@segment", SqlDbType.NVarChar);
            cmd.Parameters["@segment"].Value = "goal" + firstOfMonthDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                mtdRevenueGoal = float.Parse(dt.Rows[0]["displayName"].ToString());
            }



            //***** Monthly Pace
            double x = uc.CalculateRevenuePace(firstOfMonthDate, yesterday);
            mtdRevenuePace = float.Parse(x.ToString());




            //mtdSwiped = uc.ReturnSquareSwiped(firstOfMonthDate, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));




            //**** Last Year Monthly Revenue
            //dbConnect dcArchive = new dbConnect();
            //dcArchive.OpenArchiveConnection();
            dt = new DataTable();

            cmd = new SqlCommand();
            //cmd.Connection = dcArchive.conn;
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Parse(firstOfMonthDate).AddYears(-1);

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Parse(endOfMonthDate).AddYears(-1);

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalRev"].ToString()))
                {
                    lastyearMTDRevenue = float.Parse(dt.Rows[0]["totalRev"].ToString());
                }
            }



            ////**** Open A/R
            //GoogleSheetsConnect gc = new GoogleSheetsConnect();
            //MoveDailyWage m = new MoveDailyWage();
            //m = gc.ConnectSheetsGetOpenAR("1W4BkpxBvZ8XsOUKD14c1q6DHp5LMW1VG5nYg67o2iPs", "Junk Invoices!I1:J1");



            json = json + "\n*" + firstOfMonthDate + " - " + yesterday + "* (Goal)\n";
            json = json + "  Pace: " + mtdRevenuePace.ToString("C") + "  (" + mtdRevenueGoal.ToString("C") + ")\n";
            json = json + "      " + DateTime.Parse(firstOfMonthDate).ToString("MMMM") + " " + DateTime.Parse(firstOfMonthDate).AddYears(-1).Year + ": " + lastyearMTDRevenue.ToString("C") + "   " + ((mtdRevenuePace - lastyearMTDRevenue) / lastyearMTDRevenue).ToString("P") + "\n";
            json = json + "  AJS: " + mtdAJS.ToString("C") + "  ($400.00)\n";
            json = json + "  RPH: " + mtdRPH.ToString("C") + "  ($110.00)\n";
            json = json + "  Direct Wage: " + mtdDirectWagePercentage.ToString("P") + "  (17.5%)\n";
            json = json + "  SG&A: " + mtdDirectWagePercentage.ToString("P") + "  (13.0%)\n";
            //json = json + "    Truck: " + yesterdayTruckWagePercentage.ToString("P") + "\n";
            //json = json + "    Clerical: " + yesterdayClericalWagePercentage.ToString("P") + "\n";
            //json = json + "    Meeting: " + yesterdayMeetingWagePercentage.ToString("P") + "\n";
            //json = json + "    Training: " + yesterdayTrainingWagePercentage.ToString("P") + "\n";
            //json = json + "    Support: " + yesterdaySupportWagePercentage.ToString("P") + "\n";
            //json = json + "    Warehouse: " + yesterdayWarehouseWagePercentage.ToString("P") + "\n";
            //json = json + "    Daily Bonus: " + yesterdayBonusWagePercentage.ToString("P") + "\n";
            //json = json + "    Trainer: " + yesterdayTrainerWagePercentage.ToString("P") + "\n";
            //json = json + "  Indirect Wage: " + mtdIndirectWagePercentage.ToString("P") + "  (1.25%)\n";
            //json = json + "    Point: " + yesterdayPointWagePercentage.ToString("P") + "\n";
            //json = json + "    Fleet: " + yesterdayFleetWagePercentage.ToString("P") + "\n";
            //json = json + "  Swiped: " + mtdSwiped.ToString("P") + "  (80%)\n";
            //json = json + "  Open A/R: " + m.RPH + "\n\n";

            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1PY_Yy9NucCDBoPU2yE7vdE6yRSASgZpQLWldiFqxLEk/edit#gid=37042350'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junkweeklyscorecard").ToString().Replace("\\\\", "\\"));

        }




        public void JunkReportOutWeeklyScorecard()
        {
            int firstLine = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string json = "";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages mdw = new MoveDailyWages();

            mdw = gc.ConnectSheetsWeeklyScorecard("1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw", "GJScorecard");

            foreach (MoveDailyWage row in mdw.DailyWages)
            {
                if (firstLine == 0)
                {
                    firstLine = 1;
                    json = "{'text': '*WEEKLY SCORECARD*\n\n*" + row.RouteNumber + "*  Responsible:  Item - Actual (Goal)\n\n";
                }
                else
                {
                    if (row.RouteNumber.ToString().Trim() == "")
                    {
                        json = json + "   " + row.CrewLead + ":   " + row.RPH + " - *MISSING*  (" + row.RouteDate + ")\n";
                    }
                    else
                    {
                        json = json + "   " + row.CrewLead + ":   " + row.RPH + " - " + row.RouteNumber + "  (" + row.RouteDate + ")\n";
                    }
                }
            }

            json = json + "\n______________________________________________________\n\n";
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw/edit#gid=1987097629'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junkweeklyscorecard").ToString().Replace("\\\\", "\\"));
        }

        public void JunkReportOutVisaTransactions()
        {
            string startMonthDate = "";
            float OverallTotal = 0;
            float ThirtyDayTotal = 0;
            float SevenDayTotal = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();


            //today is the 1st
            startMonthDate = DateTime.Now.Month + "/1/" + DateTime.Now.Year;
            if (DateTime.Now.Day == 1)
            {
                if (DateTime.Now.Month == 1)
                {
                    //January 1st
                    startMonthDate = DateTime.Now.AddMonths(-1).Month + " /1/" + DateTime.Now.AddYears(-1).Year;
                }
                else
                {
                    startMonthDate = DateTime.Now.AddMonths(-1).Month + " /1/" + DateTime.Now.Year;
                }
            }

            //testing for 1st of the month
            //startMonthDate = DateTime.Now.AddMonths(-1).Month + "/1/" + DateTime.Now.Year;

            string json = "{'text': '*VISA Transactions (last 7 days)*\n";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages mdw = new MoveDailyWages();
            MoveDailyWages trans7Day = new MoveDailyWages();
            MoveDailyWages trans30Day = new MoveDailyWages();

            mdw = gc.ConnectSheetsVISACardholders("1DLvwZENrZYSW7gRYLuDGP-6KmmAyvrsbEvMnet10M9M");

            //get the cardholders and cycle through each
            foreach (MoveDailyWage row in mdw.DailyWages)
            {
                SevenDayTotal = 0;
                ThirtyDayTotal = 0;

                //get their transactions for MTD
                trans30Day = gc.ConnectSheetsVISATransactions("1DLvwZENrZYSW7gRYLuDGP-6KmmAyvrsbEvMnet10M9M", row.CrewLead, Convert.ToDateTime(startMonthDate), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
                if (trans30Day.DailyWages.Count > 0)
                {
                    foreach (MoveDailyWage thirtyday in trans30Day.DailyWages)
                    {
                        ThirtyDayTotal = ThirtyDayTotal + float.Parse(thirtyday.RouteNumber.ToString().Replace("$", ""));
                    }

                    //output their name
                    json = json + "   " + row.CrewLead + "\n";

                    //get their transactions for the last 7 days
                    trans7Day = gc.ConnectSheetsVISATransactions("1DLvwZENrZYSW7gRYLuDGP-6KmmAyvrsbEvMnet10M9M", row.CrewLead, Convert.ToDateTime(DateTime.Now.AddDays(-7).ToShortDateString()), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
                    foreach (MoveDailyWage transRow in trans7Day.DailyWages)
                    {
                        SevenDayTotal = SevenDayTotal + float.Parse(transRow.RouteNumber.ToString().Replace("$", ""));
                        json = json + "      " + transRow.RouteDate + "  " + transRow.RouteNumber + "  " + transRow.CrewLead + "\n";
                    }

                    //output the totals
                    json = json + "\n   7 Day Total:   " + SevenDayTotal.ToString("C");
                    json = json + "\n     MTD Total:   " + ThirtyDayTotal.ToString("C") + "\n";
                    json = json + "______________________________________________________\n\n";


                    OverallTotal = OverallTotal + ThirtyDayTotal;
                }



            }




            json = json + "\n   Overall Total:   " + OverallTotal.ToString("C") + "\n\n";
            json = json + "______________________________________________________\n\n";
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1DLvwZENrZYSW7gRYLuDGP-6KmmAyvrsbEvMnet10M9M'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junkvisatransactions").ToString().Replace("\\\\", "\\"));


        }


        public void JunkCheckUniformInventory()
        {
            string message = "";
            ConfigSettings cs = new ConfigSettings();
            List<UniformData> localList = new List<UniformData>();

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            localList = gc.RetrieveUniformReorderData(cs.ReturnConfigSetting("NightlyRouteToSlack", "uniformSheetID").ToString(), "gj");

            foreach(UniformData ud in localList)
            {
                message = message + ud.ItemName + ":  in stock (" + ud.ItemStock + ")   -   reorder point (" + ud.ItemReorder + ")  -  amount to reorder (" + ud.ItemReorderAmount + ")\n";
            }


            SlackClient slack = new SlackClient();
            string json = "{'text': '*Junk Uniform Reorder Notification*" + "\n" + message + "'}";
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_junk_checkuniforms").ToString().Replace("\\\\", "\\"));

        }





        public void JunkSquareSwipePercentage()
        {
            float totalJobs = 0;
            float swipedJobs = 0;
            float keyedJobs = 0;

            float keyedRev = 0;
            float keyedFees = 0;

            float swipeChipRev = 0;
            float swipeChipFees = 0;

            float lostMoney = 0;

            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareSwiped> localList = new List<SquareSwiped>();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveSquareSwiped(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));

            foreach (SquareSwiped x in localList)
            {
                totalJobs = totalJobs + 1;
                switch(x.ProcessingType.ToString().ToLower())
                {

                    case "keyed":
                        keyedJobs = keyedJobs + 1;
                        keyedRev = keyedRev + float.Parse(x.Revenue);
                        keyedFees = keyedFees + float.Parse(x.ProceesingFee);
                        break;

                    default:
                        swipedJobs = swipedJobs + 1;
                        swipeChipRev = swipeChipRev + float.Parse(x.Revenue);
                        swipeChipFees = swipeChipFees + float.Parse(x.ProceesingFee);
                        break;

                }
            }

            lostMoney = keyedRev * float.Parse(".0055");

            string json = "";
            SlackClient slack = new SlackClient();

            json = "{'text': 'Swiped Revenue: " + String.Format("{0:C}", swipeChipRev) + "\nSwiped Fees (2.15%): " + String.Format("{0:C}", swipeChipFees) + "\nSwiped Jobs: " + swipedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", swipedJobs / totalJobs) + " (Goal:  80%)\n\nKeyed Revenue: " + String.Format("{0:C}", keyedRev) + "\nKeyed Fees (2.7%): " + String.Format("{0:C}", keyedFees) + "\nKeyed Jobs: " + keyedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", keyedJobs / totalJobs) + "\n\nMoney Given Away (annualized): " + String.Format("{0:C}", lostMoney) + "   (" + String.Format("{0:C}", lostMoney * 365) + ")'}";
            //if ((swipedJobs / totalJobs) >= .8)
            //{
            //    json = "{'text': 'Swiped Revenue: " + String.Format("{0:C}", swipeChipRev) + "\nSwiped Fees (2.15%): " + String.Format("{0:C}", swipeChipFees) + "\nSwiped Jobs: " + swipedJobs + " of " + totalJobs + " - *" + string.Format("{0:P2}", swipedJobs / totalJobs) + "*\n\n*Nice work hitting the 80% goal .... Keep driving swiped transactions!!*\n\nKeyed Revenue: " + String.Format("{0:C}", keyedRev) + "\nKeyed Fees (2.7%): " + String.Format("{0:C}", keyedFees) + "\nKeyed Jobs: " + keyedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", keyedJobs / totalJobs) + "\n\nMoney Given Away (annualized): " + String.Format("{0:C}", lostMoney) + "   (" + String.Format("{0:C}", lostMoney * 365) + ")'}";
            //}
            //else
            //{
            //    json = "{'text': 'Swiped Revenue: " + String.Format("{0:C}", swipeChipRev) + "\nSwiped Fees (2.15%): " + String.Format("{0:C}", swipeChipFees) + "\nSwiped Jobs: " + swipedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", swipedJobs / totalJobs) + "\n\n*Come on ... this team can do better!*\n\nKeyed Revenue: " + String.Format("{0:C}", keyedRev) + "\nKeyed Fees (2.7%): " + String.Format("{0:C}", keyedFees) + "\nKeyed Jobs: " + keyedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", keyedJobs / totalJobs) + "\n\n*Money Given Away (annualized): " + String.Format("{0:C}", lostMoney) + "   (" + String.Format("{0:C}", lostMoney * 365) + ")*'}";
            //}
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunksquareswipe").ToString().Replace("\\\\", "\\"));



        }


        public void RunCheckChecklists()
        {
            ConfigSettings cs = new ConfigSettings();
            iAuditorClient iac = new iAuditorClient();
            RingCentralClient rcc = new RingCentralClient();
            SlackClient sc = new SlackClient();
            AuditChecklist localAD = new AuditChecklist();

            //**** get authorization code from iAuditor API ****
            UtilityClass uc = new UtilityClass();
            string authToken = "";
            authToken = iac.GetiAuditorAPIToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "iAuditorJunkAPIToken"));

            //template_171616F0F1684CA9839281F12FF0FF61 = Beginning of Day / DOT Checklist
            List<Audit> auditList = uc.RetrieveiAuditorAudits(authToken, "template_171616F0F1684CA9839281F12FF0FF61");

            for (int i = 0; i < auditList.Count; i++)
            {
                //**** check for already saved to database ... means it is already been processed ****
                dbConnect dc = new dbConnect();
                dc.OpenMessageConnection();

                DataTable dt = new DataTable();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetCheckMorningChecklist";

                cmd.Parameters.Add("@auditID", SqlDbType.NVarChar);
                cmd.Parameters["@auditID"].Value = auditList[i].audit_id;

                SqlDataAdapter ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);
                dc.CloseConnection();


                if (dt.Rows.Count <= 0)
                {

                    //  **** get Date and Route Number from Audit
                    localAD = uc.RetrieveJunkAuditChecklist(auditList[i].audit_id, authToken, true);

                    //*********** UPDATE INFORMATION IN DATABASE ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateMorningChecklist";
                    cmd.Parameters.AddWithValue("@bizUnit", "gj");
                    cmd.Parameters.AddWithValue("@dateSent", localAD.auditDate);
                    cmd.Parameters.AddWithValue("@route", localAD.auditRoute);
                    cmd.Parameters.AddWithValue("@auditID", localAD.auditID);

                    cmd.ExecuteNonQuery();

                }

            }



            //template_A1B81FB0CA144DF3A868B902CA08785D = End of Day Check Checklist
            auditList = uc.RetrieveiAuditorAudits(authToken, "template_A1B81FB0CA144DF3A868B902CA08785D");

            for (int i = 0; i < auditList.Count; i++)
            {
                //**** check for already saved to database ... means it is already been processed ****
                dbConnect dc = new dbConnect();
                dc.OpenMessageConnection();

                DataTable dt = new DataTable();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetCheckEveningChecklist";

                cmd.Parameters.Add("@auditID", SqlDbType.NVarChar);
                cmd.Parameters["@auditID"].Value = auditList[i].audit_id;

                SqlDataAdapter ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);
                dc.CloseConnection();


                if (dt.Rows.Count <= 0)
                {

                    //  **** get Date and Route Number from Audit
                    localAD = uc.RetrieveJunkAuditChecklist(auditList[i].audit_id, authToken, false);

                    //*********** UPDATE INFORMATION IN DATABASE ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateEveningChecklist";
                    cmd.Parameters.AddWithValue("@bizUnit", "gj");
                    cmd.Parameters.AddWithValue("@dateSent", localAD.auditDate);
                    cmd.Parameters.AddWithValue("@route", localAD.auditRoute);
                    cmd.Parameters.AddWithValue("@auditID", localAD.auditID);

                    cmd.ExecuteNonQuery();

                }

            }

        }

        
        public void DownloadSquareTransactions()
        {

            //go out to Square and get the transactions from today ...
            List<SquareData> localList = new List<SquareData>();
            ConfigSettings cs = new ConfigSettings();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveOvernightSquareTransactions(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));

            try
            {
                //go out and grab the RM tips for each job id
                List<SquareData> updatedList = new List<SquareData>();
                updatedList = ReturnUpdatedListWithRMTips(localList);

                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "114T2kCA174JbPTuW_kd_x1LF-N7eEAVhpBkj6NHmGbE";

                //used for testing
                //string spreadSheetId = "1hU1GxDYrkFzakFs8uSFEFkc96Xp5Gnv-p3pUmFjNGzg";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateJunkSquareTransactions(spreadSheetId, updatedList);


            }
            catch (WebException ex)
            {

            }



        }

        public void DownloadSquareTransactionsWest()
        {

            //go out to Square and get the transactions from today ...
            List<SquareData> localList = new List<SquareData>();
            ConfigSettings cs = new ConfigSettings();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveOvernightSquareTransactions(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkWestSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));

            try
            {
                //go out and grab the RM tips for each job id
                List<SquareData> updatedList = new List<SquareData>();
                updatedList = ReturnUpdatedListWithRMTips(localList);

                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "16omenlHluHGTvRFmNEcF_Y8jq4hfYLFcxj9Nk1Wfcpg";

                //used for testing
                //string spreadSheetId = "1hU1GxDYrkFzakFs8uSFEFkc96Xp5Gnv-p3pUmFjNGzg";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateJunkSquareTransactions(spreadSheetId, updatedList);


            }
            catch (WebException ex)
            {

            }



        }



        public List<SquareData> ReturnUpdatedListWithRMTips(List<SquareData> localList)
        {
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand();
            List<SquareData> returnList = new List<SquareData>();

            foreach(SquareData x in localList)
            {
                if(!(x.JobID.ToString().Trim() == ""))
                {
                    dt = new DataTable();
                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spGetJobByJobID";

                    cmd.Parameters.Add("@ticketID", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketID"].Value = x.JobID;

                    SqlDataAdapter ds = new SqlDataAdapter(cmd);
                    ds.Fill(dt);

                    if (!(dt.Rows.Count == 0))
                    {
                        x.RMTip = dt.Rows[0][9].ToString();
                        x.RMLink = "http://routes.tcjunkhauling.com/RouteEntry.aspx?rid=" + dt.Rows[0][1].ToString();
                        x.RMMatches = "FALSE";

                        //Revenue
                        if (x.Revenue == dt.Rows[0]["ticketRevenue"].ToString().Trim() &
                                x.MattressCount == dt.Rows[0]["ticketMattressesNumber"].ToString().Trim() &
                                x.MattressValue == dt.Rows[0]["ticketMattressesValue"].ToString().Trim() &
                                x.TVLargeCount == dt.Rows[0]["ticketTVLargeNumber"].ToString().Trim() &
                                x.TVLargeValue == dt.Rows[0]["ticketTVLargeValue"].ToString().Trim() &
                                x.TVSmallCount == dt.Rows[0]["ticketTVSmallNumber"].ToString().Trim() &
                                x.TVSmallValue == dt.Rows[0]["ticketTVSmallValue"].ToString().Trim() &
                                x.TiresCount == dt.Rows[0]["ticketTiresNumber"].ToString().Trim() &
                                x.TiresValue == dt.Rows[0]["ticketTiresValue"].ToString().Trim() &
                                x.Tax == dt.Rows[0]["ticketTax"].ToString().Trim() &
                                (float.Parse(x.Discount) * -1).ToString().Trim() == dt.Rows[0]["ticketDiscount"].ToString().Trim() &
                                x.Tip == dt.Rows[0]["ticketTips"].ToString().Trim()
                                )
                        {
                            x.RMMatches = "TRUE";
                        }
                    }
                    else
                    {
                        x.RMTip = "0";
                        x.RMLink = "";
                        x.RMMatches = "FALSE";
                    }

                }


            }

            dc.CloseConnection();

            return localList;

        }


        public void ImportSquareTransactions(string locationID)
        {


            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareData> localList = new List<SquareData>();
            SquareClient sc = new SquareClient();

            //localList = sc.RetrieveSquareTransactions(false, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));
            localList = sc.RetrieveSquareTransactions(false, locationID, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareAccessToken"));



            //import said transactions if not already found ...
            string routeNumber = "";
            string routeID = "";
            int routeNum = 0;
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand();

            foreach (SquareData x in localList)
            {
                routeNumber = "99";
                if (x.Route.ToString().ToLower().IndexOf("truck") > -1)
                {
                    routeNumber = x.Route.ToString().ToLower().Replace("truck", "").Trim();
                }
                if (!int.TryParse(routeNumber, out routeNum))
                {
                    routeNumber = "99";
                }



                dt = new DataTable();
                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetRouteByDateNumber";

                cmd.Parameters.Add("@routeDate", SqlDbType.NVarChar);
                cmd.Parameters["@routeDate"].Value = x.RouteDate;

                cmd.Parameters.Add("@routeNumber", SqlDbType.NVarChar);
                cmd.Parameters["@routeNumber"].Value = routeNumber;

                SqlDataAdapter ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);

                if (dt.Rows.Count == 0)
                {
                    //route not found ... so much create route first
                    

                    routeID = Guid.NewGuid().ToString();


                    //insert route
                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertRouteByDate";

                    cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                    cmd.Parameters["@id"].Value = routeID;

                    cmd.Parameters.Add("@routeDate", SqlDbType.NVarChar);
                    cmd.Parameters["@routeDate"].Value = x.RouteDate;

                    cmd.Parameters.Add("@routeNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@routeNumber"].Value = routeNumber;

                    cmd.ExecuteNonQuery();


                    ////create psRouteItems
                    //cmd = new SqlCommand();
                    //cmd.Connection = dc.conn;
                    //cmd.CommandType = CommandType.StoredProcedure;
                    //cmd.CommandText = "spGetProfitShareItems";

                    //ds = new SqlDataAdapter(cmd);
                    //ds.Fill(dt);

                    //foreach (DataRow dr in dt.Rows)
                    //{
                    //    cmd = new SqlCommand();
                    //    cmd.Connection = dc.conn;
                    //    cmd.CommandType = CommandType.StoredProcedure;
                    //    cmd.CommandText = "spInsertPSRouteItem";

                    //    cmd.Parameters.Add("@itemID", SqlDbType.NVarChar);
                    //    cmd.Parameters["@itemID"].Value = dr["id"].ToString();

                    //    cmd.Parameters.Add("@routeID", SqlDbType.NVarChar);
                    //    cmd.Parameters["@routeID"].Value = routeID;

                    //    cmd.Parameters.Add("@itemValue", SqlDbType.NVarChar);
                    //    cmd.Parameters["@itemValue"].Value = "N";

                    //    cmd.ExecuteNonQuery();

                    //}


                    ////create route marketing items



                    ////create route top 5
                    //cmd = new SqlCommand();
                    //cmd.Connection = dc.conn;
                    //cmd.CommandType = CommandType.StoredProcedure;
                    //cmd.CommandText = "spInsertTop5";

                    //cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                    //cmd.Parameters["@id"].Value = Guid.NewGuid().ToString();

                    //cmd.Parameters.Add("@routeID", SqlDbType.NVarChar);
                    //cmd.Parameters["@routeID"].Value = routeID;

                    //cmd.Parameters.Add("@goal1", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal1"].Value = "<div class=\"divTop5Goal1\"><strong><u>EVERY</u></strong> JOB EXPECTATIONS<br /><ul><li>1) CALL AHEAD 15-30 MINS BEFORE SCHEDULED TIME</li><li>2) ARRIVE ON-TIME </li><li>3) UP-SELL </li><li>4) CLEANUP </li><li>5) ASK FOR REFERRALS </li><li>6) ASK FOR A SIGN IN THE YARD </li><li>7) 20-30 DOORHANGERS</li></div>";

                    //cmd.Parameters.Add("@goal1value", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal1value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal2", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal2"].Value = "Revenue Goal:  ";

                    //cmd.Parameters.Add("@goal2value", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal2value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal3", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal3"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal3value", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal3value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal4", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal4"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal4value", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal4value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal5", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal5"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@goal5value", SqlDbType.NVarChar);
                    //cmd.Parameters["@goal5value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave1", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave1"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave1value", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave1value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave2", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave2"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave2value", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave2value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave3", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave3"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@wave3value", SqlDbType.NVarChar);
                    //cmd.Parameters["@wave3value"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@trucklocation", SqlDbType.NVarChar);
                    //cmd.Parameters["@trucklocation"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@truckboxspace", SqlDbType.NVarChar);
                    //cmd.Parameters["@truckboxspace"].Value = DBNull.Value;

                    //cmd.Parameters.Add("@trucknotes", SqlDbType.NVarChar);
                    //cmd.Parameters["@trucknotes"].Value = DBNull.Value;

                    //cmd.ExecuteNonQuery();


                }





                //if route is or is not found, check to see if job id is found



                //get route ID
                dt = new DataTable();
                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetRouteByDateNumber";

                cmd.Parameters.Add("@routeDate", SqlDbType.NVarChar);
                cmd.Parameters["@routeDate"].Value = x.RouteDate;

                cmd.Parameters.Add("@routeNumber", SqlDbType.NVarChar);
                cmd.Parameters["@routeNumber"].Value = routeNumber;

                ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);

                routeID = dt.Rows[0]["id"].ToString();





                dt = new DataTable();
                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetJobByJobID";

                cmd.Parameters.Add("@ticketID", SqlDbType.NVarChar);
                cmd.Parameters["@ticketID"].Value = x.JobID;

                ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);

                if (dt.Rows.Count == 0)
                {
                    //job not found ... so import 
                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertJob";

                    cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                    cmd.Parameters["@id"].Value = Guid.NewGuid().ToString();

                    cmd.Parameters.Add("@routeID", SqlDbType.NVarChar);
                    cmd.Parameters["@routeID"].Value = routeID;

                    cmd.Parameters.Add("@jobID", SqlDbType.NVarChar);
                    cmd.Parameters["@jobID"].Value = x.JobID;

                    cmd.Parameters.Add("@jobType", SqlDbType.NVarChar);
                    cmd.Parameters["@jobType"].Value = x.JobType;

                    cmd.Parameters.Add("@ticketNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketNumber"].Value = "";

                    cmd.Parameters.Add("@ticketRevenue", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketRevenue"].Value = x.Revenue;

                    cmd.Parameters.Add("@ticketDiscount", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketDiscount"].Value = x.Discount.Replace("-", "");

                    cmd.Parameters.Add("@ticketCEC", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketCEC"].Value = 0;

                    cmd.Parameters.Add("@ticketNet", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketNet"].Value = x.NetRevenue;

                    cmd.Parameters.Add("@ticketTips", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTips"].Value = x.Tip;

                    cmd.Parameters.Add("@sign", SqlDbType.NVarChar);
                    cmd.Parameters["@sign"].Value = "N";

                    cmd.Parameters.Add("@payMethod", SqlDbType.NVarChar);
                    cmd.Parameters["@payMethod"].Value = x.PayMethod;

                    cmd.Parameters.Add("@doorHangers", SqlDbType.NVarChar);
                    cmd.Parameters["@doorHangers"].Value = "N";

                    cmd.Parameters.Add("@ticketDetails", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketDetails"].Value = "imported";

                    cmd.Parameters.Add("@ticketTax", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTax"].Value = x.Tax;

                    cmd.Parameters.Add("@ticketTVLargeNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTVLargeNumber"].Value = x.TVLargeCount;

                    cmd.Parameters.Add("@ticketTVLargeValue", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTVLargeValue"].Value = x.TVLargeValue;

                    cmd.Parameters.Add("@ticketTVSmallNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTVSmallNumber"].Value = x.TVSmallCount;

                    cmd.Parameters.Add("@ticketTVSmallValue", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTVSmallValue"].Value = x.TVSmallValue;

                    cmd.Parameters.Add("@ticketTiresNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTiresNumber"].Value = x.TiresCount;

                    cmd.Parameters.Add("@ticketTiresValue", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketTiresValue"].Value = x.TiresValue;

                    cmd.Parameters.Add("@ticketMattressesNumber", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketMattressesNumber"].Value = x.MattressCount;

                    cmd.Parameters.Add("@ticketMattressesValue", SqlDbType.NVarChar);
                    cmd.Parameters["@ticketMattressesValue"].Value = x.MattressValue;

                    cmd.Parameters.Add("@reasonNotConvert", SqlDbType.NVarChar);
                    cmd.Parameters["@reasonNotConvert"].Value = "**Select if applicable";

                    cmd.Parameters.Add("@estimator", SqlDbType.NVarChar);
                    cmd.Parameters["@estimator"].Value = "Not Applicable";

                    cmd.ExecuteNonQuery();

                }



            }

            dc.CloseConnection();


        }




        public void RunDailyMinDumps()
        {
            SlackClient sc = new SlackClient();
            string output = "";

            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetMinDumps";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            foreach (DataRow dr in dt.Rows)
            {
                output = output + "Route " + dr["routeNumber"].ToString().Trim() + ";  " + dr["dumpLocation"].ToString().Trim() + ": " + String.Format("{0:F}", dr["ticketTons"]) + " tons.   " + String.Format("{0:C}", dr["ticketTotal"]) + "\n";
            }

            try
            {
                if (output == "")
                {
                    output = "None from yesterday";
                }

                ConfigSettings cs = new ConfigSettings();
                string json = "{'text': '*Min Dumps*\n" + output + "'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkdumplocations").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                ConfigSettings cs = new ConfigSettings();
                sc.SendMessage("{'text': 'There was an error with slack_rundailymindumps today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }


        }

        public void RunDumpLocation()
        {
            SlackClient sc = new SlackClient();
            string output = "";
            double total = 0;
            double count = 0;
            double overallTotal = 0;
            double overallCount = 0;

            double hercTotal = 0;
            double bpTotal = 0;

            string startDate = "1/1/";
            string localYear = DateTime.Today.Year.ToString();
            if (DateTime.Today.Month.ToString() == "1")
            {
                localYear = DateTime.Today.AddDays(-1).Year.ToString();
            }
            startDate = startDate + localYear;

            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spReport_GetDumpsByLocationTotal";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                overallTotal = Convert.ToDouble(dt.Rows[0]["routeValueSum"].ToString());
                overallCount = Convert.ToDouble(dt.Rows[0]["routeNumberCount"].ToString());

                DataTable dtInner = new DataTable();

                cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spReport_GetDumpsByLocationOverall";

                cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
                cmd.Parameters["@startDate"].Value = startDate;

                cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
                cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

                ds = new SqlDataAdapter(cmd);
                ds.Fill(dtInner);
                dc.CloseConnection();

                foreach (DataRow drow in dtInner.Rows)
                {
                    count = float.Parse(drow["routeNumberCount"].ToString());
                    total = float.Parse(drow["routeValueSum"].ToString());
                    if (drow["dumpLocation"].ToString().Trim().ToLower().Contains("hennepin energy resource"))
                    {
                        hercTotal = total;
                    }
                    else if (drow["dumpLocation"].ToString().Trim().ToLower().Contains("brooklyn park transfer"))
                    {
                        bpTotal = total;
                    }
                    output = output + drow["dumpLocation"].ToString().Trim() + ": " + count + " (" + String.Format("{0:P}", count / overallCount) + ") - " + String.Format("{0:C}", total) + " (" + String.Format("{0:P}", total / overallTotal) + ")\n";
                }

            }

            int dayOfYear = DateTime.Now.DayOfYear;


            output = output + "\n";
            output = output + "HERC Tons: " + String.Format("{0:F}", hercTotal / 70) + " (Pace - " + String.Format("{0:F}", ((hercTotal / 70) / dayOfYear) * 365) + ")\n";
            output = output + "BP Tons: " + String.Format("{0:F}", bpTotal / 70) + " (Pace - " + String.Format("{0:F}", ((bpTotal / 70) / dayOfYear) * 365) + ")\n";
            output = output + "Total: " + String.Format("{0:F}", bpTotal / 70 + hercTotal / 70) + " (Pace - " + String.Format("{0:F}", (((bpTotal / 70) / dayOfYear) * 365) + ((hercTotal / 70) / dayOfYear) * 365) + ")";

            try
            {
                ConfigSettings cs = new ConfigSettings();
                string json = "{'text': '*Dumps By Location: YTD " + DateTime.Today.AddDays(-1).ToShortDateString() + "*\n" + output + "'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkdumplocations").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                ConfigSettings cs = new ConfigSettings();
                sc.SendMessage("{'text': 'There was an error with slack_rundumplocation today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }

        public void RunOverallDailyRouteRevenue()
        {
            SlackClient sc = new SlackClient();

            //******************* DAILY OVERALL ROUTE DATA *******************

            double totalRev = 0;
            double countRev = 0;
            double totalWithRev = 0;
            double totalJobs = 0;
            double totalDumps = 0;
            double totalWages = 0;
            double totalHours = 0;


            //**** Total Revenue ****
            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                totalRev = Convert.ToDouble(dt.Rows[0]["totalRev"].ToString());
                countRev = Convert.ToDouble(dt.Rows[0]["countRev"].ToString());
            }


            //**** Total Wages ****
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetWagesByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                totalWages = Convert.ToDouble(dt.Rows[0]["totalPay"].ToString());
                totalHours = Convert.ToDouble(dt.Rows[0]["totalHours"].ToString());
            }


            //**** Total OSC ****
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetOSCByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                totalWithRev = Convert.ToDouble(dt.Rows[0]["totalWithRev"].ToString());
                totalJobs = Convert.ToDouble(dt.Rows[0]["totalJobs"].ToString());
            }



            //**** Total Dumps ****
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetDumpsByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = DateTime.Today.AddDays(-1).ToShortDateString();

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!(dt.Rows[0]["countDumps"].ToString() == "0"))
                {
                    totalDumps = Convert.ToDouble(dt.Rows[0]["totalDumps"].ToString());
                }
            }



            //**** Send the message ****
            ConfigSettings cs = new ConfigSettings();
            try
            {
                string json = "{'text': 'All Routes:   " + DateTime.Today.AddDays(-1).ToShortDateString() + "\nRevenue: " + String.Format("{0:C}", totalRev) + "\nAJS: " + String.Format("{0:C}", (totalRev / countRev)) + " - " + countRev + " Jobs" + "\nRPH: " + String.Format("{0:C}", (totalRev / totalHours)) + " - All hours: " + totalHours + "\nDumps: " + String.Format("{0:P2}", (totalDumps / totalRev)) + " - " + String.Format("{0:C}", totalDumps) + "\nOSC: " + String.Format("{0:P2}", (totalWithRev / totalJobs)) + " - " + totalWithRev + " of " + totalJobs + " with revenue" + "\nWages: " + String.Format("{0:P2}", (totalWages / totalRev)) + " - " + String.Format("{0:C}", totalWages) + "'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkdailyroutedata").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_gotjunkdailyroutedata today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }



        public void RunDailySchedule(bool testing)
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            string rcTokenURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL").ToString().Replace("\\\\", "\\");
            string rcTokenData = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataJunk").ToString().Replace("\\\\", "\\");
            string rcTokenAuth = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth").ToString().Replace("\\\\", "\\");
            string rcJunkURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSJunkURL").ToString().Replace("\\\\", "\\");
            string rcJunkPhone = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSJunkPhone").ToString().Replace("\\\\", "\\");
            string rcImageLocation = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralMMSImageLocation").ToString().Replace("\\\\", "\\").Replace("\\", "\\\\");
            string scDailySchedule = cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkdailyschedule").ToString().Replace("\\\\", "\\");

            //get schedule/employee information
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();
            List<JunkDailySchedule> jds = new List<JunkDailySchedule>();

            List<JunkDailySchedule> jdsWest = new List<JunkDailySchedule>();
            jdsWest = gsc.ConnectSheetsJunkDailySchedule("West Schedule!", "L1:N100");
            foreach (JunkDailySchedule js in jdsWest)
            {
                JunkDailySchedule jsTemp = new JunkDailySchedule();
                jsTemp.EmpName = js.EmpName;
                jsTemp.EmpPhone = js.EmpPhone;
                jsTemp.EmpStartTime = js.EmpStartTime;
                jsTemp.EmpWorkType = js.EmpWorkType;

                jds.Add(jsTemp);
            }


            List<JunkDailySchedule> jdsEast = new List<JunkDailySchedule>();
            jdsEast = gsc.ConnectSheetsJunkDailySchedule("East Schedule!", "L1:N100");
            foreach (JunkDailySchedule js in jdsEast)
            {
                JunkDailySchedule jsTemp = new JunkDailySchedule();
                jsTemp.EmpName = js.EmpName;
                jsTemp.EmpPhone = js.EmpPhone;
                jsTemp.EmpStartTime = js.EmpStartTime;
                jsTemp.EmpWorkType = js.EmpWorkType;

                jds.Add(jsTemp);
            }



            //get ring central token
            RingCentralClient rc = new RingCentralClient();
            string accessToken = rc.GetRingCentralToken(rcTokenURL, rcTokenData, rcTokenAuth);


            //send text message
            string scMessage = "{'text': 'Messages Sent: \n";
            foreach (JunkDailySchedule js in jds)
            {
                if (testing)
                {
                    //testing send to myself
                    rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "16122813895", "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + " on route " + js.EmpWorkType + ".", "junk_truck.png");
                }
                else
                {
                    //Production
                    rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "1" + js.EmpPhone, "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + " on route " + js.EmpWorkType + ".", "junk_truck.png");
                }

                scMessage = scMessage + "   " + js.EmpName + " - " + js.EmpPhone + "  " + js.EmpStartTime + " " + js.EmpWorkType + "\n";

                //insert into database
                dbConnect dc = new dbConnect();
                dc.OpenMessageConnection();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spInsertSent";
                cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                cmd.Parameters.AddWithValue("@bizUnit", "gj");
                cmd.Parameters.AddWithValue("@dateSent", DateTime.Now.ToString());
                cmd.Parameters.AddWithValue("@phoneTo", "1" + js.EmpPhone);
                cmd.Parameters.AddWithValue("@truckTeamLead", "289f6b3c-0392-4efc-9eb7-becb35200e19");      //Avery
                cmd.Parameters.AddWithValue("@mediaURL", "junk_truck.png");
                cmd.Parameters.AddWithValue("@messageType", "dailyschedule");
                cmd.Parameters.AddWithValue("@message", "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + " on route " + js.EmpWorkType + ".");
                cmd.Parameters.AddWithValue("@sentby", "289f6b3c-0392-4efc-9eb7-becb35200e19");     //Avery

                cmd.ExecuteNonQuery();


                if (!js.EmpWorkType.ToLower().Contains("team lead") & !js.EmpWorkType.ToLower().Contains("on-call") & !js.EmpWorkType.ToLower().Contains("point") & !string.IsNullOrEmpty(js.EmpWorkType.Trim()))
                {
                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertRouteChecklists";
                    cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@bizUnit", "gj");
                    cmd.Parameters.AddWithValue("@dateSent", DateTime.Now.AddDays(1).ToShortDateString());
                    cmd.Parameters.AddWithValue("@route", js.EmpWorkType);
                    cmd.Parameters.AddWithValue("@routeTeam", js.EmpName);
                    cmd.Parameters.AddWithValue("@morningChecklist", "N");
                    cmd.Parameters.AddWithValue("@morningAuditID", "");
                    cmd.Parameters.AddWithValue("@eveningChecklist", "N");
                    cmd.Parameters.AddWithValue("@eveningAuditID", "");

                    cmd.ExecuteNonQuery();
                }


            }
            scMessage = scMessage + "'}";

            //send to slack
            sc.SendMessage(scMessage, scDailySchedule);

        }



        public void RunDailyWagePercentage()
        {
            UtilityClass uc = new UtilityClass();
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();

            //******************* DAILY WAGE PERCENTAGE *******************
            double totalRev = 0;
            double goalRev = 0;
            double countRev = 0;
            double totalDirectLabor = 0;
            double totalIndirectLabor = 0;
            double totalDumps = 0;
            double pace = 0;
            double dayRevenueNeeded = 0;
            double messagesSent = 0;
            double messagesOpened = 0;
            string startDate = cc.GetStartDate();
            string endDate = DateTime.Today.AddDays(-1).ToShortDateString();

            dbConnect dc = new dbConnect();
            dc.OpenConnection();

            //**** Total Revenue
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetRevByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            //cmd.Parameters["@startDate"].Value = "1/1/2019";
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            //cmd.Parameters["@endDate"].Value = "3/31/2019";
            cmd.Parameters["@endDate"].Value = endDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalRev"].ToString()))
                {
                    totalRev = Convert.ToDouble(dt.Rows[0]["totalRev"].ToString());
                    countRev = Convert.ToDouble(dt.Rows[0]["countRev"].ToString());
                }
            }
            

            //**** Goal Revenue
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetDropDownListData";

            cmd.Parameters.Add("@segment", SqlDbType.NVarChar);
            cmd.Parameters["@segment"].Value = "goal" + startDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                goalRev = Convert.ToDouble(dt.Rows[0]["displayName"].ToString());
            }


            //**** Total Wages
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomated_GetWagesByDates";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["totalDirectLabor"].ToString()))
                {
                    totalDirectLabor = Convert.ToDouble(dt.Rows[0]["totalDirectLabor"].ToString());
                    totalIndirectLabor = Convert.ToDouble(dt.Rows[0]["totalIndirectLabor"].ToString());
                }
            }


            ////**** Total Dumps ****
            //dt = new DataTable();

            //cmd = new SqlCommand();
            //cmd.Connection = dc.conn;
            //cmd.CommandType = CommandType.StoredProcedure;
            //cmd.CommandText = "spAutomated_GetDumpsByDates";

            //cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            //cmd.Parameters["@startDate"].Value = startDate;

            //cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            //cmd.Parameters["@endDate"].Value = endDate;

            //ds = new SqlDataAdapter(cmd);
            //ds.Fill(dt);
            //dc.CloseConnection();

            //if (dt.Rows.Count > 0)
            //{
            //    if (!(dt.Rows[0]["countDumps"].ToString() == "0"))
            //    {
            //        totalDumps = Convert.ToDouble(dt.Rows[0]["totalDumps"].ToString());
            //    }
            //}


            //***** Day Revenue Needed
            DateTime start = Convert.ToDateTime(startDate);
            DateTime end = Convert.ToDateTime(endDate);
            dayRevenueNeeded = cc.CalculateDailyRevenueNeed(totalRev, goalRev, start, end);


            //***** Monthly Pace
            pace = uc.CalculateRevenuePace(startDate, endDate);


            dc = new dbConnect();
            dc.OpenMessageConnection();

            //**** Messages Sent
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomatedGetWelcomeTextSent";

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = "gjnd";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                messagesSent = Convert.ToDouble(dt.Rows[0]["opportunityIDCount"].ToString());
            }

            //**** Messages Opened
            dt = new DataTable();

            cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spAutomatedGetWelcomeTextOpened";

            cmd.Parameters.Add("@bizUnit", SqlDbType.NVarChar);
            cmd.Parameters["@bizUnit"].Value = "gjnd";

            cmd.Parameters.Add("@startDate", SqlDbType.NVarChar);
            cmd.Parameters["@startDate"].Value = startDate;

            cmd.Parameters.Add("@endDate", SqlDbType.NVarChar);
            cmd.Parameters["@endDate"].Value = endDate;

            ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                messagesOpened = Convert.ToDouble(dt.Rows[0]["opportunityIDCount"].ToString());
            }


            //Review Requests Sent
            Review r = new Review();
            double ReviewRequestsSent = r.ReturnReviewsRequested("gjjd");

            //Google Reviews Received
            int NumberOfReviews = r.ReturnReviewsNumber("gj", start.Month.ToString(), start.Year.ToString());
            //int NumberOfReviews1 = r.ReturnReviewsNumber("gj", "1", start.Year.ToString());
            //int NumberOfReviews2 = r.ReturnReviewsNumber("gj", "2", start.Year.ToString());
            //int NumberOfReviews3 = r.ReturnReviewsNumber("gj", "3", start.Year.ToString());


            //no longer doing resale
            //float resale = 0;
            //resale = uc.RetrieveWarehouseSale();
            //totalDumps = totalDumps - resale;

            try
            {
                ConfigSettings cs = new ConfigSettings();
                //string json = "{'text': 'Date Range:  MTD " + endDate + "\nRevenue: " + String.Format("{0:C}", totalRev) + "\nJobs: " + countRev + " (" + String.Format("{0:C}", totalRev / countRev) + " AJS)" + "\nGoal: " + String.Format("{0:C}", goalRev) + "\nMonthly Pace: " + String.Format("{0:C}", pace) + "\n*Today Revenue Need: " + String.Format("{0:C}", dayRevenueNeeded) + "*\nDirect Wages: " + String.Format("{0:C}", totalDirectLabor) + " (" + String.Format("{0:P2}", (totalDirectLabor / totalRev)) + ")\nIndirect Wages: " + String.Format("{0:C}", totalIndirectLabor) + " (" + String.Format("{0:P2}", (totalIndirectLabor / totalRev)) + ")\nDumps: " + String.Format("{0:C}", totalDumps) + " (" + String.Format("{0:P2}", (totalDumps / totalRev)) + ")\nReviews Requested: " + ReviewRequestsSent + " (" + String.Format("{0:P2}", ReviewRequestsSent / countRev) + ")" + "\nReviews Received: " + NumberOfReviews + " (" + String.Format("{0:P2}", NumberOfReviews / countRev) + ")\nWelcome Texts Sent: " + messagesSent + "\nWelcome Texts Opened: " + messagesOpened + " (" + String.Format("{0:P2}", messagesOpened / messagesSent) + ")'}";
                string json = "{'text': 'Date Range:  MTD " + endDate + "\nRevenue: " + String.Format("{0:C}", totalRev) + "\nJobs: " + countRev + " (" + String.Format("{0:C}", totalRev / countRev) + " AJS)" + "\nGoal: " + String.Format("{0:C}", goalRev) + "\nMonthly Pace: " + String.Format("{0:C}", pace) + "\n*Today Revenue Need: " + String.Format("{0:C}", dayRevenueNeeded) + "*\nDirect Wages: " + String.Format("{0:C}", totalDirectLabor) + " (" + String.Format("{0:P2}", (totalDirectLabor / totalRev)) + ")\nIndirect Wages: " + String.Format("{0:C}", totalIndirectLabor) + " (" + String.Format("{0:P2}", (totalIndirectLabor / totalRev)) + ")\nReviews Requested: " + ReviewRequestsSent + " (" + String.Format("{0:P2}", ReviewRequestsSent / countRev) + ")" + "\nReviews Received: " + NumberOfReviews + " (" + String.Format("{0:P2}", NumberOfReviews / countRev) + ")\nWelcome Texts Sent: " + messagesSent + "\nWelcome Texts Opened: " + messagesOpened + " (" + String.Format("{0:P2}", messagesOpened / messagesSent) + ")'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkwages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                ConfigSettings cs = new ConfigSettings();
                sc.SendMessage("{'text': 'There was an error with slack_rundailywagepercentage today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }


        }

    }
}
