using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utilities360Wow;
using System.Configuration;
using Google.Apis.Sheets.v4;
using System.Net;
using System.Data;
using System.Data.SqlClient;

namespace NightlyRouteToSlack.Utilities
{
    public class ShackUtilities
    {

        public void ShackNPSTracker()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            int x = 0;
            string json = "{'text': '*DAILY NPS TRACKING*\n\n";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages m = new MoveDailyWages();

            //NOTE:  I can't pull from the O2E owned sheet so we created a local sheet that pulls in the O2E sheet
            //          https://docs.google.com/spreadsheets/d/1jnGmLblCOcaM-QoumpjXsGbn8sy3QdtfqBSf9lxtSQc
            m = gc.ConnectSheetsShackNPSScorecard("1phtEpvJXGmRI_Am2NTOGcR6sSL8yjeV4qesUHSO33VE");

            json = json + "*" + m.startDate + " - " + m.endDate + "*\n";
            json = json + "   Overall NPS: " + m.score + "         *Goal: " + m.goal + "*\n\n";
            json = json + "   *Tech Leads*\n\n";


            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                if (mdw.EndDate.Trim().ToLower() == "stop")
                {
                    json = json + "   " + mdw.StartDate;
                }
                else
                {
                    json = json + "   " + mdw.StartDate + ":   " + mdw.EndDate + "    " + mdw.CrewLead + " response(s)\n";
                }
            }
            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1jnGmLblCOcaM-QoumpjXsGbn8sy3QdtfqBSf9lxtSQc'}";

            //NPS Scorecard Slack Channel
            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shacknpsscorecard").ToString().Replace("\\\\", "\\"));

            //All Team Slack Channel
            //sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackallteam").ToString().Replace("\\\\", "\\"));







            json = "{'text': '*POP INS TRACKING*\n\n";

            //NOTE:  I can't pull from the O2E owned sheet so we created a local sheet that pulls in the O2E sheet
            //          https://docs.google.com/spreadsheets/d/1jnGmLblCOcaM-QoumpjXsGbn8sy3QdtfqBSf9lxtSQc
            m = gc.ConnectSheetsShackPopInsTracking("1phtEpvJXGmRI_Am2NTOGcR6sSL8yjeV4qesUHSO33VE");

            json = json + "*" + m.startDate + "*\n";
            json = json + "   Number of Jobs: " + m.endDate + "\n";
            json = json + "   Pop Ins Completed: " + m.goal + "\n";
            json = json + "   Average Score: " + m.avg + "\n";
            json = json + "   Pop Ins Percentage: " + m.score + "\n\n";
            json = json + "   *Tech Leads*\n";


            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                json = json + "   " + mdw.StartDate + ":   " + mdw.EndDate + "    " + mdw.CrewLead + " response(s)\n";
            }
            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1jnGmLblCOcaM-QoumpjXsGbn8sy3QdtfqBSf9lxtSQc'}";

            //NPS Scorecard Slack Channel
            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shacknpsscorecard").ToString().Replace("\\\\", "\\"));

            //All Team Slack Channel
            //sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackallteam").ToString().Replace("\\\\", "\\"));





        }



        public void ShackDailyChecklist()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            int x = 0;
            string json = "{'text': '*DAILY CHECKLIST - ";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages m = new MoveDailyWages();

            m = gc.ConnectSheetsShackDailyOpsChecklist("1ET7HXLs8q-ZDvnn1EdhjP6pwUPnbw4ZpfwaU6dcDPoc");

            json = json + m.DailyWages[0].RPH + " - " + m.DailyWages[0].WagePercentage + "*\n\n";
            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                if (x == 0)
                {
                    x = 1;
                }
                else
                {
                    json = json + "  " + mdw.RouteDate + "\n";
                }
            }
            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1ET7HXLs8q-ZDvnn1EdhjP6pwUPnbw4ZpfwaU6dcDPoc'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackdailyopschecklist").ToString().Replace("\\\\", "\\"));



            //reset sheet back to "empty"
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1ET7HXLs8q-ZDvnn1EdhjP6pwUPnbw4ZpfwaU6dcDPoc";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateShackDailyOpsChecklist(spreadSheetId);


            }
            catch (WebException ex)
            {

            }


        }



        public void ShackDailyScorecard() 
        {
            CommonClient cc = new CommonClient();
            UtilityClass uc = new UtilityClass();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            float yesterdaySwiped = 0;
            float mtdRevenueGoal = 0;
            float mtdSwiped = 0;

            string json = "{'text': '";

            string yesterday = DateTime.Today.AddDays(-1).ToShortDateString();
            string firstOfMonthDate = cc.GetStartDate();
            string ajsGoal = uc.RetrieveGoal("ssGoalAJS" + firstOfMonthDate);
            string rphGoal = uc.RetrieveGoal("ssGoalRPH" + firstOfMonthDate);

            //string yesterday = "8/21/2020";
            //string firstOfMonthDate = "8/15/2020";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWage m = new MoveDailyWage();

            //*********************** DAILY SCORECARD ***********************
            m = gc.ConnectSheetsShackDailyScorecard(DateTime.Parse(yesterday), DateTime.Parse(yesterday));

            //yesterdaySwiped = uc.ReturnSquareSwiped(yesterday, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareAccessToken"));

            json = json + "*DAILY SCORECARD*\n\n*" + yesterday + "*\n";
            json = json + "  Revenue: " + float.Parse(m.RouteDate).ToString("C") + "\n";
            json = json + "  AJS: " + float.Parse(m.WagePercentage).ToString("C") + "  (" + float.Parse(ajsGoal.Trim()).ToString("C") + ")\n";
            json = json + "  RPH: " + float.Parse(m.RPH).ToString("C") + "  (" + float.Parse(rphGoal.Trim()).ToString("C") + ")\n";
            json = json + "  Direct Wage: " + float.Parse(m.RouteNumber).ToString("P") + " (" + m.MiscItem2.ToString() + ")\n";
            //json = json + "  Payment Received: " + float.Parse(m.CrewLead).ToString("C") + "\n";
            //json = json + "  Swiped: " + yesterdaySwiped.ToString("P") + "\n";


            //*********************** MONTHLY SCORECARD ***********************
            m = gc.ConnectSheetsShackDailyScorecard(DateTime.Parse(firstOfMonthDate), DateTime.Parse(yesterday));

            //**** Revenue Pace
            ShackWagesMessageModel mwm = new ShackWagesMessageModel();
            mwm = gc.ConnectSheetsShackProductionPace();

            if (mwm.ProductionPace.IndexOf("DIV",0) > 0)
            {
                mwm.ProductionPace = "0";
            }

            if (mwm.OverallRevenueMTD.IndexOf("DIV", 0) > 0)
            {
                mwm.OverallRevenueMTD = "0";
            }

            //mtdSwiped = uc.ReturnSquareSwiped(firstOfMonthDate, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareAccessToken"));



            //**** Open A/R
            MoveDailyWage mLocal = new MoveDailyWage();
            mLocal = gc.ConnectSheetsGetOpenAR("1W4BkpxBvZ8XsOUKD14c1q6DHp5LMW1VG5nYg67o2iPs", "Shack Invoices!G1:H1");




            json = json + "\n*" + firstOfMonthDate + " - " + yesterday + "* (Goal)\n";
            json = json + "  Revenue: " + float.Parse(m.RouteDate).ToString("C") + "\n";
            json = json + "  Overall Pace: " + float.Parse(mwm.ProductionPace.Replace("$", "").Replace(" ", "")).ToString("C") + " (" + float.Parse(mwm.OverallRevenueMTD.Replace("$", "").Replace(" ", "")).ToString("C") + ")\n";
            json = json + "    Detailing Pace: " + float.Parse(mwm.DetailingRevenuePace.Replace("$", "").Replace(" ", "")).ToString("C") + " (" + float.Parse(mwm.DetailingRevenue.Replace("$", "").Replace(" ", "")).ToString("C") + ")\n";
            json = json + "    Lights Pace: " + float.Parse(mwm.LightsRevenuePace.Replace("$", "").Replace(" ", "")).ToString("C") + " (" + float.Parse(mwm.LightsRevenue.Replace("$", "").Replace(" ", "")).ToString("C") + ")\n";
            //json = json + "  Pace: " + float.Parse(mwm.ProductionPace.Replace("$", "").Replace(" ", "")).ToString("C") + " (" + mtdRevenueGoal.ToString("C") + ")\n";
            json = json + "  AJS: " + float.Parse(m.WagePercentage).ToString("C") + "  (" + float.Parse(ajsGoal.Trim()).ToString("C") + ")\n";
            json = json + "  RPH: " + float.Parse(m.RPH).ToString("C") + "  (" + float.Parse(rphGoal.Trim()).ToString("C") + ")\n";
            json = json + "  Direct Wage: " + m.MiscItem.ToString() + " (" + m.MiscItem2.ToString() + ")\n";
            json = json + "  OPS OH: " + m.MiscItem3.ToString() + " (" + m.MiscItem4.ToString() + ")\n";
            //json = json + "  Payment Received: " + float.Parse(m.CrewLead).ToString("C") + " (All of it!)\n";
            //json = json + "  Swiped: " + mtdSwiped.ToString("P") + " (80%)\n";
            json = json + "  Open A/R: " + float.Parse(mLocal.RPH.Replace("$", "").Replace("-", "")).ToString("C") + "  (minimal)\n";




            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1PY_Yy9NucCDBoPU2yE7vdE6yRSASgZpQLWldiFqxLEk/edit#gid=2058938254'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackweeklyscorecard").ToString().Replace("\\\\", "\\"));




        }



        public void ShackSendMessages()
        {
            int counter = 0;
            ConfigSettings cs = new ConfigSettings();
            RingCentralClient rcc = new RingCentralClient();

            dbConnect dc = new dbConnect();
            dc.OpenMessageConnection();

            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetMessagesToSend";
            cmd.Parameters.AddWithValue("@bizUnit", "ss");

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                //*********** GET RING CENTRAL TOKEN ***********
                string accessToken = "";
                accessToken = rcc.GetRingCentralToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataShack"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth"));

                DateTime updatedDate;
                DateTime goodDate = DateTime.Now;
                DateTime errDate = DateTime.Parse("1/1/1900");
                foreach (DataRow dr in dt.Rows)
                {
                    try
                    {
                        //*********** FOR EACH RETURNED ROW, SEND TEXT MESSAGE TO PHONE ***********
                        rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackPhone"), dr[4].ToString(), dr[5].ToString());
                        updatedDate = goodDate;

                        //*********** UPDATE ACTUAL NUMBER OF SENT **************
                        counter = counter + 1;
                    }
                    catch (Exception e)
                    {
                        updatedDate = errDate;
                    }

                    //*********** FOR EACH RETURNED ROW, UPDATE DATE SENT ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateMessagesToSend";
                    cmd.Parameters.AddWithValue("@messageID", dr[0].ToString());
                    cmd.Parameters.AddWithValue("@dateSent", updatedDate);

                    cmd.ExecuteNonQuery();
                    dc.CloseConnection();

                }

                //*********** SEND TEXT MESSAGE TO OM ***********
                rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackPhone"), cs.ReturnConfigSetting("NightlyRouteToSlack", "shackSendMessagesOMPhone"), counter + " messages sent - " + DateTime.Now.ToString());
                //rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSShackPhone"), "16122813895", counter + " messages sent - " + DateTime.Now.ToString());
            }

        }


        public void ShackReportOutWeeklyScorecard()
        {
            int firstLine = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string json = "";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages mdw = new MoveDailyWages();

            mdw = gc.ConnectSheetsWeeklyScorecard("1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw", "SSScorecard");

            foreach (MoveDailyWage row in mdw.DailyWages)
            {
                if (firstLine == 0)
                {
                    firstLine = 1;
                    json = "{'text': '*" + row.RouteNumber + "*  Responsible:  Item - Actual (Goal)\n\n";
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
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw/edit#gid=38972350'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackweeklyscorecard").ToString().Replace("\\\\", "\\"));

        }


        public void ShackReportOutVisaTransactions()
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

            mdw = gc.ConnectSheetsVISACardholders("1OILQ4ixna1UzTot7Jy5kXPvLc70RhV3e5Xz62QqGfbI");

            foreach (MoveDailyWage row in mdw.DailyWages)
            {
                SevenDayTotal = 0;
                ThirtyDayTotal = 0;

                //get their transactions for MTD
                trans30Day = gc.ConnectSheetsVISATransactions("1OILQ4ixna1UzTot7Jy5kXPvLc70RhV3e5Xz62QqGfbI", row.CrewLead, Convert.ToDateTime(startMonthDate), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
                if (trans30Day.DailyWages.Count > 0)
                {
                    foreach (MoveDailyWage thirtyday in trans30Day.DailyWages)
                    {
                        ThirtyDayTotal = ThirtyDayTotal + float.Parse(thirtyday.RouteNumber.ToString().Replace("$", ""));
                    }

                    //output their name
                    json = json + "   " + row.CrewLead + "\n";

                    //get their transactions for the last 7 days
                    trans7Day = gc.ConnectSheetsVISATransactions("1OILQ4ixna1UzTot7Jy5kXPvLc70RhV3e5Xz62QqGfbI", row.CrewLead, Convert.ToDateTime(DateTime.Now.AddDays(-7).ToShortDateString()), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
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
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1OILQ4ixna1UzTot7Jy5kXPvLc70RhV3e5Xz62QqGfbI'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackvisatransactions").ToString().Replace("\\\\", "\\"));


        }


        public void ShackCheckUniformInventory()
        {
            string message = "";
            ConfigSettings cs = new ConfigSettings();
            List<UniformData> localList = new List<UniformData>();

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            localList = gc.RetrieveUniformReorderData(cs.ReturnConfigSetting("NightlyRouteToSlack", "uniformSheetID").ToString(), "ss");

            foreach (UniformData ud in localList)
            {
                message = message + ud.ItemName + ":  in stock (" + ud.ItemStock + ")   -   reorder point (" + ud.ItemReorder + ")  -  amount to reorder (" + ud.ItemReorderAmount + ")\n";
            }


            SlackClient slack = new SlackClient();
            string json = "{'text': '*Shack Uniform Reorder Notification*" + "\n" + message + "'}";
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shack_checkuniforms").ToString().Replace("\\\\", "\\"));

        }


        public void ShackSquareSwipePercentage()
        {
            float totalJobs = 0;
            float keyedJobs = 0;
            float swipedJobs = 0;

            float keyedRev = 0;
            float keyedFees = 0;

            float swipeChipRev = 0;
            float swipeChipFees = 0;

            float lostMoney = 0;

            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareSwiped> localList = new List<SquareSwiped>();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveSquareSwiped(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareAccessToken"));

            foreach (SquareSwiped x in localList)
            {
                totalJobs = totalJobs + 1;
                switch (x.ProcessingType.ToString().ToLower())
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

            if ((swipedJobs / totalJobs) >= .8)
            {
                json = "{'text': 'Swiped Revenue: " + String.Format("{0:C}", swipeChipRev) + "\nSwiped Fees (2.15%): " + String.Format("{0:C}", swipeChipFees) + "\nSwiped Jobs: " + swipedJobs + " of " + totalJobs + " - *" + string.Format("{0:P2}", swipedJobs / totalJobs) + "*\n\n*Nice work hitting the 80% goal .... Keep driving swiped transactions!!*\n\nKeyed Revenue: " + String.Format("{0:C}", keyedRev) + "\nKeyed Fees (2.7%): " + String.Format("{0:C}", keyedFees) + "\nKeyed Jobs: " + keyedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", keyedJobs / totalJobs) + "\n\nMoney Given Away (annualized): " + String.Format("{0:C}", lostMoney) + "   (" + String.Format("{0:C}", lostMoney * 365) + ")'}";
            }
            else
            {
                json = "{'text': 'Swiped Revenue: " + String.Format("{0:C}", swipeChipRev) + "\nSwiped Fees (2.15%): " + String.Format("{0:C}", swipeChipFees) + "\nSwiped Jobs: " + swipedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", swipedJobs / totalJobs) + "\n\n*Come on ... this team can do better!*\n\nKeyed Revenue: " + String.Format("{0:C}", keyedRev) + "\nKeyed Fees (2.7%): " + String.Format("{0:C}", keyedFees) + "\nKeyed Jobs: " + keyedJobs + " of " + totalJobs + " - " + string.Format("{0:P2}", keyedJobs / totalJobs) + "\n\n*Money Given Away (annualized): " + String.Format("{0:C}", lostMoney) + "   (" + String.Format("{0:C}", lostMoney * 365) + ")*'}";
            }
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shacksquareswipe").ToString().Replace("\\\\", "\\"));



        }


        public void ShackShineUpdateOTAwareness()
        {
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateShackOTAwareness();

            }
            catch (WebException ex)
            {

            }

        }



        public void DownloadSquareTransactions()
        {

            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareDataShack> localList = new List<SquareDataShack>();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveSquareTransactionsShack(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackSquareAccessToken"));

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                //2019 sheet
                //reader.UpdateShackSquareTransactions("1BsHtO4sV8LGw2ad5C6Ysmt2EHw6aIRvjYjGObCOOSaU", localList);

                //2020 sheet
                reader.UpdateShackSquareTransactions("13DS58gdeylKMAwjxqEqaclMTob_UG5LROe4JfaJfDK0", localList);

            }
            catch (WebException ex)
            {

            }



        }





        public void SendShackSalesPercentageToSlack()
        {
            SlackClient sc = new SlackClient();
            ShackSales ss = new ShackSales();
            ConfigSettings cs = new ConfigSettings();

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            ss = gc.ConnectSheetsShackSales();

            try
            {
                string json = "{'text': '30 Day Date Range: " + ss.StartDate30Day + " - " + ss.EndDate30Day + "\nDetailing Estimates / Converted / Percentage: " + ss.DetailingEstimates30Day + " / " + ss.DetailingConverted30Day + " / " + ss.DetailingPercentage30Day + "\nLights Estimates / Converted / Percentage: " + ss.LightsEstimates30Day + " / " + ss.LightsConverted30Day + " / " + ss.LightsPercentage30Day + "\n\n7 Day Date Range: " + ss.StartDate7Day + " - " + ss.EndDate7Day + "\nDetailing Estimates / Converted / Percentage: " + ss.DetailingEstimates7Day + " / " + ss.DetailingConverted7Day + " / " + ss.DetailingPercentage7Day + "\nLights Estimates / Converted / Percentage: " + ss.LightsEstimates7Day + " / " + ss.LightsConverted7Day + " / " + ss.LightsPercentage7Day + "'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shacksales").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_sendshacksalespercentage today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }
        }

        public void RunDailyRouteWagesShack()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            try
            {
                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                ShackDailyWages sdw = new ShackDailyWages();

                sdw = gc.ConnectSheetsShackDailyWage();

                string json = "{'text': '";
                foreach (ShackDailyWage m in sdw.DailyWages)
                {
                    json = json + "Date: " + m.RouteDate + "\nRoute: " + m.RouteNumber + "\nRevenue: " + m.Revenue + "\n";


                    double result;

                    result = Convert.ToDouble(m.WagePercentage.Replace("%", ""));
                    if (result / 100 >= .30)
                    {
                        //wage percentage is above 25%
                        json = json + "*Wage %: " + m.WagePercentage + "*\n";
                    }
                    else
                    {
                        json = json + "Wage %: " + m.WagePercentage + "\n";
                    }


                    result = Convert.ToDouble(m.RPH.Replace("$", ""));
                    if (result < 50)
                    {
                        //wage percentage is above 25%
                        json = json + "*RPH: " + m.RPH + "*\n\n";
                    }
                    else
                    {
                        json = json + "RPH: " + m.RPH + "\n\n";
                    }


                }
                json = json + "'}";

                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackwages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_shackwages today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }

        public void RunDailyWagePercentageShack()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            try
            {
                string startDate = cc.GetStartDate();
                DateTime start = Convert.ToDateTime(startDate);

                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                ShackWagesMessageModel mwm = new ShackWagesMessageModel();

                mwm = gc.ConnectSheetsShack();

                //Review Requests Sent
                Review r = new Review();
                double ReviewRequestsSent = r.ReturnReviewsRequested("ssjd");

                //Google Reviews Received
                int NumberOfReviews = r.ReturnReviewsNumber("ss", start.Month.ToString(), start.Year.ToString());

                string json = "{'text': 'Date Range:  " + mwm.StartDate + " - " + mwm.EndDate +
                    "\nLabor Revenue: " + String.Format("{0:C}", mwm.Revenue) + 
                    "\n\nOverall Revenue (Pace): " + String.Format("{0:C}", mwm.OverallRevenueMTD) + "   (" + String.Format("{0:C}", mwm.ProductionPace) + ")" +
                    "\n   Lights Revenue (Pace): " + String.Format("{0:C}", mwm.LightsRevenue) + "   (" + String.Format("{0:C}", mwm.LightsRevenuePace) + ")" +
                    "\n   Detailing Revenue (Pace): " + String.Format("{0:C}", mwm.DetailingRevenue) + "   (" + String.Format("{0:C}", mwm.DetailingRevenuePace) + ")" +
                    "\n\nSales Pace: " + String.Format("{0:C}", mwm.SalesPace) +
                    "\n   Lights Sales (Pace): " + String.Format("{0:C}", mwm.LightsSales) + "   (" + String.Format("{0:C}", mwm.LightsSalesPace) + ")" +
                    "\n   Detailing Sales (Pace): " + String.Format("{0:C}", mwm.DetailingSales) + "   (" + String.Format("{0:C}", mwm.DetailingSalesPace) + ")" +
                    "\n\nNet Wage %: " + mwm.NetWagePercentage + " (" + mwm.GoalWagePercentage + ")" + 
                    "\nMtg. Tech Wage %: " + mwm.MktgWagePercentage + " (" + mwm.MktgGoalWagePercentage + ")" + 
                    "\nSales Wage %: " + mwm.SalesWagePercentage + " (" + mwm.SalesGoalWagePercentage + ")" + 
                    "\nOH Wage %: " + mwm.OHWagePercentage + " (" + mwm.OHGoalWagePercentage + ")" + 
                    "\nAdmin Wage %: " + mwm.AdminWagePercentage + " (" + mwm.AdminGoalWagePercentage + ")" +
                    "\n\nOverall OH %: " + mwm.OverallOHPercentage + " (" + mwm.OverallOHGoalPercentage + ")" +
                    "\n\nReviews Requested: " + ReviewRequestsSent + 
                    "\nReviews Received: " + NumberOfReviews + "'}";


                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_shackwages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_rundailywagepercentageshack today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }


    }
}
