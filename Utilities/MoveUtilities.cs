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
using Newtonsoft.Json.Linq;

namespace NightlyRouteToSlack.Utilities
{
    public class MoveUtilities
    {

        public void MoveDamageSlackOutput()
        {
            float totalDamage = 0;
            float totalService = 0;
            float totalTotal = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            GoogleSheetsConnect gc = new GoogleSheetsConnect();


            MoveDailyWages m = new MoveDailyWages();
            m = gc.ConnectSheetsMoveDamage("1Vcavfwan9FbZ0O2En4G6FkQ_6lL_3J1XhF1rtVlgwV0");

            string json = "{'text': '*CUSTOMER FOLLOW UP*\n*" + m.DailyWages[0].StartDate + " - " + m.DailyWages[0].EndDate + "*\n  Crew Lead | Damage $$ | Service $$ | Total $$\n\n";

            foreach (MoveDailyWage mLocal in m.DailyWages)
            {
                json = json + "  " + mLocal.CrewLead + ":     " + mLocal.MiscItem + " | " + mLocal.MiscItem2 + " | " + mLocal.MiscItem3 + "\n";
                totalDamage = totalDamage + float.Parse(mLocal.MiscItem.Replace("$", "").Replace(",", ""));
                totalService = totalService + float.Parse(mLocal.MiscItem2.Replace("$", "").Replace(",", ""));
                totalTotal = totalTotal + float.Parse(mLocal.MiscItem3.Replace("$", "").Replace(",", ""));
            }

            json = json + "\n  Totals:     " + totalDamage.ToString("C") + " | " + totalService.ToString("C") + " | " + totalTotal.ToString("C") + "\n";


            json = json + "\n______________________________________________________\n";
            json = json + "\n  Reference Doc : https://docs.google.com/spreadsheets/d/1Vcavfwan9FbZ0O2En4G6FkQ_6lL_3J1XhF1rtVlgwV0/edit#gid=49840077'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedailydamageupdate").ToString().Replace("\\\\", "\\"));


        }

        public void MoveCheckJakeEnteringHours()
        {
            ConfigSettings cs = new ConfigSettings();
            RingCentralClient rcc = new RingCentralClient();
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();

            if (gsc.DidJakeWorkTodayWithoutTimeCardEntry())
            {
                //*********** GET RING CENTRAL TOKEN ***********
                string accessToken = "";
                accessToken = rcc.GetRingCentralToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataMove"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth"));

                //*********** SEND TEXT MESSAGE ***********
                //rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMoveURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMovePhone"), "16122813895", "Hey Jake ... reminder to enter your time card hours for today!");
                rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMoveURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMovePhone"), "19529946994", "Hey Jake ... reminder to enter your time card hours for today!");
            }




        }

        public void MoveSalesCenterConversion()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            GoogleSheetsConnect gc = new GoogleSheetsConnect();


            //*********************** GET CONVERSION DATA ***********************
            MoveDailyWages m = new MoveDailyWages();
            m = gc.ConnectSheetsSalesCenterConversion("1T_htIBFPsNe2xCgXFvgUdSec3s_-nupKs4S204HfTr4");

            //30 day calls -> CrewLead
            //30 day booked -> WagePercentage
            //30 day conv -> RPH
            //30 day ose/booked -> RouteDate
            //30 day conv -> RouteNumber
            //estimator -> MiscItem
            //date range calls -> MiscItem2
            //date range booked ->MiscItem3
            //date range conv ->MiscItem4
            //date range ose/booked ->MiscItem5
            //date range conv ->MiscItem6

            string json = "{'text': '*SALES CENTER CONVERSION*\n\n";

            foreach (MoveDailyWage mLocal in m.DailyWages)
            {
                json = json + "  *" + mLocal.MiscItem + "*\n";
                json = json + "     " + mLocal.StartDate + "-" + mLocal.EndDate + "\n";
                json = json + "        Calls: " + mLocal.MiscItem2 + "\n";
                json = json + "        Jobs Booked: " + mLocal.MiscItem3 + "\n";
                json = json + "        Gross Conversion: " + mLocal.MiscItem4 + "\n";
                json = json + "        OSEs Booked: " + mLocal.MiscItem5 + "\n";
                json = json + "        Net Conversion: " + mLocal.MiscItems6 + "\n";
                json = json + "     Rolling 30 days :\n";
                json = json + "        Calls: " + mLocal.CrewLead + "\n";
                json = json + "        Jobs Booked: " + mLocal.WagePercentage + "\n";
                json = json + "        Gross Conversion: " + mLocal.RPH + "\n";
                json = json + "        OSEs Booked: " + mLocal.RouteDate + "\n";
                json = json + "        Net Conversion: " + mLocal.RouteNumber + "\n\n";
            }

            json = json + "\n______________________________________________________\n";
            json = json + "\n  Reference Doc : https://docs.google.com/spreadsheets/d/1T_htIBFPsNe2xCgXFvgUdSec3s_-nupKs4S204HfTr4/edit#gid=201842141'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movesalescenterconversion").ToString().Replace("\\\\", "\\"));

        }


        public void MoveSalesCenterWages()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWage q = new MoveDailyWage();

            string wagePercentage = "0";
            string firstOfMonthDate = cc.GetStartDate();
            string yesterday = DateTime.Today.ToShortDateString();

            //*********************** GET MTD REVENUE ***********************
            //m = gc.ConnectSheetsMoveDailyScorecard(DateTime.Parse(firstOfMonthDate), DateTime.Parse(yesterday));
            //m.RouteDate = "101464.40";

            //*********************** GET MTD SALES CENTER WAGES ***********************
            //q = gc.ConnectSheetsMoveSalesCenterWages(DateTime.Parse(firstOfMonthDate), DateTime.Parse(yesterday));
            //q.RouteDate = "1487.33";

            //**** DETERMINE WAGE PERCENTAGE ****
            //wagePercentage = (float.Parse(q.RouteDate.Replace("$", "").Replace(" ", "")) / float.Parse(m.RouteDate.Replace("$", "").Replace(" ", ""))).ToString();



            m = gc.ConnectSheetsSalesWages();

            //**** OUTPUT ****
            string json = "{'text': '*YMM SALES*\n";
            json = json + "  Wage % :  " + m.MiscItem.ToString() + " (" + m.MiscItem2.ToString() + ")\n";
            json = json + "  Over/Under % :  " + m.MiscItem3.ToString() + "\n";
            json = json + "  Over/Under $ :  " + m.MiscItem4.ToString() + "\n\n";

            json = json + "*SS - SALES*\n";
            json = json + "  Wage % :  " + m.MiscItem5.ToString() + " (" + m.MiscItems6.ToString() + ")\n";
            json = json + "  Over/Under % :  " + m.CrewLead.ToString() + "\n";
            json = json + "  Over/Under $ :  " + m.EndDate.ToString() + "\n";

            json = json + "\n______________________________________________________\n";
            json = json + "\n  Reference Docs:";
            json = json + "\n    https://docs.google.com/spreadsheets/d/1MTwqtKym02eP6xHX3OmmbnfNtxG9CocS_YCTrQ_rqiM/edit#gid=1236509418'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movesalescenterwages").ToString().Replace("\\\\", "\\"));

        }



        public void MoveDailyChecklist()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            int x = 0;
            string json = "{'text': '*DAILY CHECKLIST - ";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages m = new MoveDailyWages();

            m = gc.ConnectSheetsDailyChecklist("1rMFYFIfiuKe6si-10CTHXR1tUMbNMYnD7nJ__9V-PBU");

            json = json + m.DailyWages[0].RPH + "*\n\n";
            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                if (x == 0)
                {
                    x = 1;
                }
                else 
                {
                    json = json + "  " + mdw.RouteDate + ": " + mdw.RouteNumber + "\n";
                }
            }
            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1rMFYFIfiuKe6si-10CTHXR1tUMbNMYnD7nJ__9V-PBU'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedailyopschecklist").ToString().Replace("\\\\", "\\"));



            //reset sheet back to "empty"
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1rMFYFIfiuKe6si-10CTHXR1tUMbNMYnD7nJ__9V-PBU";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateMoveDailyOpsChecklist(spreadSheetId, m);


            }
            catch (WebException ex)
            {

            }



        }


        public void MoveDailyScorecard() 
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
            float yesterdayIndirectWage = 0;
            float yesterdayIndirectWagePercentage = 0;
            float yesterdaySwiped = 0;
            float mtdRevenue = 0;
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

            string json = "{'text': '";

            string yesterday = DateTime.Today.AddDays(-1).ToShortDateString();
            string firstOfMonthDate = cc.GetStartDate();

            //string yesterday = "3/17/2020";
            //string firstOfMonthDate = "3/1/2020";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWage m = new MoveDailyWage();

            //*********************** DAILY SCORECARD ***********************
            m = gc.ConnectSheetsMoveDailyScorecard(DateTime.Parse(yesterday), DateTime.Parse(yesterday));

            yesterdaySwiped = uc.ReturnSquareSwiped(yesterday, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareAccessToken"));

            json = json + "*DAILY SCORECARD*\n\n*" + yesterday + "*\n";
            json = json + "  Revenue: " + float.Parse(m.RouteDate).ToString("C") + "\n";
            json = json + "  AJS: " + float.Parse(m.WagePercentage).ToString("C") + "  ($1,200.00)\n";
            json = json + "  RPH: " + float.Parse(m.RPH).ToString("C") + " ($80.00)\n";
            json = json + "  Direct Wage: " + float.Parse(m.RouteNumber).ToString("P") + " (" + m.MiscItem2.ToString() + ")\n";
            json = json + "  Swiped: " + yesterdaySwiped.ToString("P") + "\n";


            //*********************** MONTHLY SCORECARD ***********************
            m = gc.ConnectSheetsMoveDailyScorecard(DateTime.Parse(firstOfMonthDate), DateTime.Parse(yesterday));

            //**** Revenue Pace
            MoveWagesMessageModel mwm = new MoveWagesMessageModel();
            mwm = gc.ConnectSheets();

            //**** Monthly Goal Revenue
            dbConnect dc = new dbConnect();
            dc.OpenConnection();
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetDropDownListData";

            cmd.Parameters.Add("@segment", SqlDbType.NVarChar);
            cmd.Parameters["@segment"].Value = "ymGoal" + firstOfMonthDate;

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                mtdRevenueGoal = float.Parse(dt.Rows[0]["displayName"].ToString());
            }

            mtdSwiped = uc.ReturnSquareSwiped(firstOfMonthDate, yesterday, cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareAccessToken"));


            //**** Open A/R
            MoveDailyWage mLocal = new MoveDailyWage();
            mLocal = gc.ConnectSheetsGetOpenAR("1W4BkpxBvZ8XsOUKD14c1q6DHp5LMW1VG5nYg67o2iPs", "Move Invoices!G1:H1");


            json = json + "\n*" + firstOfMonthDate + " - " + yesterday + "* (Goal)\n";
            json = json + "  Revenue: " + float.Parse(m.RouteDate).ToString("C") + "\n";
            json = json + "  Pace: " + float.Parse(mwm.RevenuePace.Replace("$","").Replace(" ","")).ToString("C") + " (" + mtdRevenueGoal.ToString("C") + ")\n";
            json = json + "  AJS: " + float.Parse(m.WagePercentage).ToString("C") + " ($1,200.00)\n";
            json = json + "  RPH: " + float.Parse(m.RPH).ToString("C") + " ($80.00)\n";
            json = json + "  Direct Wage: " + m.MiscItem.ToString() + " (" + m.MiscItem2.ToString() + ")\n";
            json = json + "  SG&A: " + m.MiscItem3.ToString() + " (" + m.MiscItem4.ToString() + ")\n";
            json = json + "  Swiped: " + mtdSwiped.ToString("P") + " (80%)\n";
            json = json + "  Open A/R: " + float.Parse(mLocal.RPH.Replace("$", "").Replace("-", "")).ToString("C") + "  (minimal)\n";







            //json = json + "\n*" + firstOfMonthDate + " - " + yesterday + "*\n";
            //json = json + "  Damage $ / %: TBD\n";
            //json = json + "  Service $ / %: TBD\n";
            //json = json + "  Pop-ins # / %: TBD\n";

            json = json + "\n______________________________________________________\n";
            json = json + "\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1PY_Yy9NucCDBoPU2yE7vdE6yRSASgZpQLWldiFqxLEk/edit#gid=1064029565'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_moveweeklyscorecard").ToString().Replace("\\\\", "\\"));



        }



        public void MoveReportOutWeeklyScorecard()
        {
            int firstLine = 0;
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            string json = "";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            MoveDailyWages mdw = new MoveDailyWages();

            mdw = gc.ConnectSheetsWeeklyScorecard("1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw", "YMScorecard");

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
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1MpjJf4RoyDPL_Re2E5E45V35ebTHorMQeKObImxZiYw/edit#gid=493682636'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_moveweeklyscorecard").ToString().Replace("\\\\", "\\"));

        }


        public void MoveReportOutVisaTransactions()
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

            mdw = gc.ConnectSheetsVISACardholders("1UmcUrlCFx9-6Sae-J5Q22RgrbuoZ2YKf_dYSrVpXw18");

            foreach(MoveDailyWage row in mdw.DailyWages)
            {
                SevenDayTotal = 0;
                ThirtyDayTotal = 0;

                //get their transactions for MTD
                trans30Day = gc.ConnectSheetsVISATransactions("1UmcUrlCFx9-6Sae-J5Q22RgrbuoZ2YKf_dYSrVpXw18", row.CrewLead, Convert.ToDateTime(startMonthDate), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
                if (trans30Day.DailyWages.Count > 0)
                {
                    foreach (MoveDailyWage thirtyday in trans30Day.DailyWages)
                    {
                        ThirtyDayTotal = ThirtyDayTotal + float.Parse(thirtyday.RouteNumber.ToString().Replace("$", ""));
                    }

                    //output their name
                    json = json + "   " + row.CrewLead + "\n";

                    //get their transactions for the last 7 days
                    trans7Day = gc.ConnectSheetsVISATransactions("1UmcUrlCFx9-6Sae-J5Q22RgrbuoZ2YKf_dYSrVpXw18", row.CrewLead, Convert.ToDateTime(DateTime.Now.AddDays(-7).ToShortDateString()), Convert.ToDateTime(DateTime.Now.AddDays(-1).ToShortDateString()));
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
            json = json + "\n\n   Reference Doc:   https://docs.google.com/spreadsheets/d/1UmcUrlCFx9-6Sae-J5Q22RgrbuoZ2YKf_dYSrVpXw18'}";

            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movevisatransactions").ToString().Replace("\\\\", "\\"));


        }


        public void MoveReportOutOverheadWages()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            try
            {
                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                MoveWagesMessageModel mwm = new MoveWagesMessageModel();

                mwm = gc.ConnectSheetsOverheadWages();

                //get RevenuePace = Gross Overhead
                //get NetWagePercentage = Move Hours Cost Reduction
                //get GoalWagePercentage = Junk/Shack Cost Reduction
                //get OHWagePercentage = Net Overhead
                //get OHGoalWagePercentage = Cash Flow Impact

                string overUnder = "Over";
                if (float.Parse(mwm.OverheadGoalResults.Replace("%", "")) < 0)
                {
                    overUnder = "Under";
                }

                string json = "{'text': 'Date Range:  " + mwm.StartDate + " - " + mwm.EndDate + "\n90% of Projected Revenue: " +
                    String.Format("{0:C}", mwm.Revenue) + "\nGross Overhead:  " + String.Format("{0:C}", mwm.RevenuePace) + 
                    " (" + mwm.GrossOverheadPercentage + ")\nMoving Hours Cost Reduction: " + String.Format("{0:C}", mwm.NetWagePercentage) + "\nJunk/Shack Cost Reduction: " +
                    String.Format("{0:C}", mwm.GoalWagePercentage) + "\nNet Overhead: " + String.Format("{0:C}", mwm.OHWagePercentage) +
                    " (" + mwm.NetOverheadPercentage + ")\nGoal: " + mwm.OverheadGoal + "\n" + mwm.OverheadGoalResults + " " + overUnder + " Budget\nCash Flow Impact:  " + String.Format("{0:C}", mwm.OHGoalWagePercentage) + "\n'}";

                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movewages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_rundailywagepercentagemove today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }

        public void MoveCheckUniformInventory()
        {
            string message = "";
            ConfigSettings cs = new ConfigSettings();
            List<UniformData> localList = new List<UniformData>();

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            localList = gc.RetrieveUniformReorderData(cs.ReturnConfigSetting("NightlyRouteToSlack", "uniformSheetID").ToString(), "ym");

            foreach (UniformData ud in localList)
            {
                message = message + ud.ItemName + ":  in stock (" + ud.ItemStock + ")   -   reorder point (" + ud.ItemReorder + ")  -  amount to reorder (" + ud.ItemReorderAmount + ")\n";
            }


            SlackClient slack = new SlackClient();
            string json = "{'text': '*Move Uniform Reorder Notification*" + "\n" + message + "'}";
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_move_checkuniforms").ToString().Replace("\\\\", "\\"));

        }


        public void MoveSquareSwipePercentage()
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

            localList = sc.RetrieveSquareSwiped(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareAccessToken"));

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
            slack.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movesquareswipe").ToString().Replace("\\\\", "\\"));



        }






        public void DownloadDailyTips()
        {
            //declare local variables
            string accessToken = "";
            SquareData sd = new SquareData();
            List<SquareData> sdList = new List<SquareData>();
            SquareClient sc = new SquareClient();
            ConfigSettings cs = new ConfigSettings();

            //set local variables
            accessToken = cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareAccessToken");

            //retrieve square payments
            List<Newtonsoft.Json.Linq.JObject> localList = new List<Newtonsoft.Json.Linq.JObject>();
            localList = sc.RetrieveSquarePaymentsReceived(true, accessToken);

            //read through payments, set list of data
            foreach (JObject x in localList)
            {
                sd = new SquareData();

                sd.RouteDate = x["created_at"].ToString().Substring(0, x["created_at"].ToString().IndexOf("/202") + 5);
                sd.TotalCollected = (float.Parse(x["total_money"]["amount"].ToString()) / 100).ToString();
                sd.Tax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                sd.Tip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();
                sd.TaxableRevenue = "0";
                sd.Total = (float.Parse(sd.TotalCollected) - float.Parse(sd.Tip) - float.Parse(sd.Tax)).ToString();

                foreach (JToken tender in x.SelectToken("tender"))
                {
                    if (tender["type"].ToString().ToLower() == "credit_card")
                    {
                        sd.PaymentType = tender["card_brand"].ToString().Replace("MASTER_CARD", "Mastercard").Replace("AMERICAN_EXPRESS", "AMEX");
                        sd.PanSuffix = tender["pan_suffix"].ToString();
                        sd.SwipeOrKeyed = tender["entry_method"].ToString();
                    }
                }

                foreach (JToken tender in x.SelectToken("itemizations"))
                {
                    if (tender["name"].ToString().ToLower() == "custom amount")
                    {
                        try
                        {
                            sd.JobID = tender["notes"].ToString();
                        }
                        catch
                        {

                        }
                    }
                }

                sdList.Add(sd);
            }

            //update list of payments to Tips worksheet
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);

                //2020
                reader.UpdateMoveTipTransactions("1vAw3-dZaDxrgA4hg4-soTAeFgUfrYQvcf1A1YItFR4k", sdList);

            }
            catch (WebException ex)
            {

            }


        }




        public void DownloadSquareTransactions()
        {

            //go out to Square and get the transactions from today ...
            ConfigSettings cs = new ConfigSettings();
            List<SquareDataMove> localList = new List<SquareDataMove>();
            SquareClient sc = new SquareClient();

            localList = sc.RetrieveSquareTransactionsMove(true, cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareLocationID"), cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveSquareAccessToken"));

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                //2019
                //reader.UpdateMoveSquareTransactions("18RJUxAyj_b0Sa7qklfo3HEIrdg1_pcTkvZ8aLN5xbZg", localList);

                //2020
                reader.UpdateMoveSquareTransactions("1vAw3-dZaDxrgA4hg4-soTAeFgUfrYQvcf1A1YItFR4k", localList);

            }
            catch (WebException ex)
            {

            }



        }





        public void SendMorningMeetingTimeToSlack()
        {
            SlackClient sc = new SlackClient();
            string meetingTime = "";

            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            meetingTime = gc.ConnectSheetsMoveMorningMeetingTime();

            try
            {
                ConfigSettings cs = new ConfigSettings();
                string json = "{'text': 'Tomorrow Morning Meeting Time: " + meetingTime + "'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_youmoveme").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                ConfigSettings cs = new ConfigSettings();
                sc.SendMessage("{'text': 'There was an error with slack_sendmorningmeetingtime today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_youmoveme").ToString().Replace("\\\\", "\\"));
            }

        }


        public void RunMoveBagDropROI()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            GoogleSheetsConnect gc = new GoogleSheetsConnect();

            string output = gc.ConnectSheetsMoveBagDropROI();

            string json = "{'text': '" + output + "'}";
            sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_ymmbagdrop").ToString().Replace("\\\\", "\\"));

        }


        public void RunUpdateReviewsReceived()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            string startDate = cc.GetStartDate();
            DateTime start = Convert.ToDateTime(startDate);

            //Google Reviews Received
            Review r = new Review();
            int NumberOfReviews = r.ReturnReviewsNumber("ym", start.Month.ToString(), start.Year.ToString());

            //Update Google Sheet
            GoogleSheetsConnect gc = new GoogleSheetsConnect();
            gc.UpdateYouMoveMeReviewsForecast(NumberOfReviews);

        }


        public void RunCheckJobDoneChecklist(string sendPhone, string templateID)
        {
            ConfigSettings cs = new ConfigSettings();
            iAuditorClient iac = new iAuditorClient();
            RingCentralClient rcc = new RingCentralClient();
            SlackClient sc = new SlackClient();
            bool sendMsg = false;
            string messageToText = "Job checklists completed: \\n";
            string messageToSlack = "";
            string messageToSlackDamage = "";
            AuditData localAD = new AuditData();

            //**** get authorization code from iAuditor API ****
            UtilityClass uc = new UtilityClass();
            string authToken = "";
            authToken = iac.GetMoveiAuditorAPIToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "iAuditorMoveAPIToken"));

            //**** get list of last 10 completed "job checklist" audits .... "template_id":"template_77e3bc7c2beb465fa4f8a433ecbfdc27"} = Job Checklist ****
            List<Audit> auditList = uc.RetrieveiAuditorAudits(authToken, templateID);

            for (int i = 0; i < auditList.Count; i++)
            {
                //**** check for already saved to database ... means it is already been processed ****
                dbConnect dc = new dbConnect();
                dc.OpenMessageConnection();

                DataTable dt = new DataTable();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spGetMoveJobDone";

                cmd.Parameters.Add("@auditID", SqlDbType.NVarChar);
                cmd.Parameters["@auditID"].Value = auditList[i].audit_id;

                SqlDataAdapter ds = new SqlDataAdapter(cmd);
                ds.Fill(dt);
                dc.CloseConnection();


                if (dt.Rows.Count <= 0)
                {
                    sendMsg = true;

                    //  **** get Customer Name and Damage information from Audit
                    localAD = uc.RetrieveiAuditorAuditCustomer(auditList[i].audit_id, authToken);

                    messageToText = messageToText + "ID: " + localAD.workOrderID + " - " + localAD.customerName + "\\n";
                    messageToSlack = messageToSlack + "ID: " + localAD.workOrderID + " - " + localAD.customerName + "\n";

                    //*********** SEND INFORMATION TO DATABASE ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertMoveJobDone";
                    cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@dateImported", DateTime.Now.ToString());
                    cmd.Parameters.AddWithValue("@auditID", auditList[i].audit_id);
                    cmd.Parameters.AddWithValue("@workOrderID", localAD.workOrderID);
                    cmd.Parameters.AddWithValue("@customerName", localAD.customerName);
                    cmd.Parameters.AddWithValue("@message", "You Move Me job checklist has been marked complete for Work Order ID: " + localAD.workOrderID + "; Customer: " + localAD.customerName);

                    cmd.ExecuteNonQuery();




                    if (localAD.damageHappenedLoad.Trim().ToLower() != "n/a" || localAD.damageHappenedUnload.Trim().ToLower() != "n/a")
                    {
                        messageToSlackDamage = messageToSlackDamage + "ID: " + localAD.workOrderID + " - " + localAD.customerName + "\nDamage (Load): " + localAD.damageHappenedLoad + "\nPrevention (Load): " + localAD.damagePreventionLoad + "\nDamage (Unload): " + localAD.damageHappenedUnload + "\nPrevention (Unload): " + localAD.damagePreventionUnload + "\n\n";

                        //*********** SEND INFORMATION TO DATABASE ***********
                        cmd = new SqlCommand();
                        cmd.Connection = dc.conn;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "spInsertMoveDamagePrevention";
                        cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                        cmd.Parameters.AddWithValue("@dateImported", DateTime.Now.ToString());
                        cmd.Parameters.AddWithValue("@auditID", auditList[i].audit_id);
                        cmd.Parameters.AddWithValue("@workOrderID", localAD.workOrderID);
                        cmd.Parameters.AddWithValue("@damageHappened", localAD.damageHappenedLoad);
                        cmd.Parameters.AddWithValue("@preventionTip", localAD.damagePreventionLoad);

                        cmd.ExecuteNonQuery();

                        cmd = new SqlCommand();
                        cmd.Connection = dc.conn;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "spInsertMoveDamagePrevention";
                        cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                        cmd.Parameters.AddWithValue("@dateImported", DateTime.Now.ToString());
                        cmd.Parameters.AddWithValue("@auditID", auditList[i].audit_id);
                        cmd.Parameters.AddWithValue("@workOrderID", localAD.workOrderID);
                        cmd.Parameters.AddWithValue("@damageHappened", localAD.damageHappenedUnload);
                        cmd.Parameters.AddWithValue("@preventionTip", localAD.damagePreventionUnload);

                        cmd.ExecuteNonQuery();

                    }


                }

            }


            if (sendMsg)
            {

                //*********** GET RING CENTRAL TOKEN ***********
                string accessToken = "";
                accessToken = rcc.GetRingCentralToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataMove"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth"));

                //*********** SEND TEXT MESSAGE ***********
                rcc.SendSMSRingCentral(accessToken, cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMoveURL"), cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMovePhone"), sendPhone, messageToText);

                //*********** SEND INFORMATION TO SLACK ***********
                try
                {
                    string json = "{'text': '" + messageToSlack + "'}";
                    sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movejobdone").ToString().Replace("\\\\", "\\"));

                    if (messageToSlackDamage != "")
                    {
                        json = "{'text': '" + messageToSlackDamage + "'}";
                        sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedamage").ToString().Replace("\\\\", "\\"));
                    }
                }
                catch
                {
                    sc.SendMessage("{'text': 'There was an error with slack_movejobdone today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
                }
            }


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
            authToken = iac.GetMoveiAuditorAPIToken(cs.ReturnConfigSetting("NightlyRouteToSlack", "iAuditorMoveAPIToken"));

            //template_5fd4ec7f5cd64a6c8863e6d399677df5 = UPDATED AM Checklist
            List<Audit> auditList = uc.RetrieveiAuditorAudits(authToken, "template_5fd4ec7f5cd64a6c8863e6d399677df5");

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
                    localAD = uc.RetrieveAuditChecklist(auditList[i].audit_id, authToken);

                    //*********** UPDATE INFORMATION IN DATABASE ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateMorningChecklist";
                    cmd.Parameters.AddWithValue("@bizUnit", "ym");
                    cmd.Parameters.AddWithValue("@dateSent", localAD.auditDate);
                    cmd.Parameters.AddWithValue("@route", localAD.auditRoute.Replace("Route", "RT"));
                    cmd.Parameters.AddWithValue("@auditID", localAD.auditID);

                    cmd.ExecuteNonQuery();

                }

            }



            //template_f3f2d50289e54ed4af110f853360a5f5 = UPDATED PM Checklist
            auditList = uc.RetrieveiAuditorAudits(authToken, "template_f3f2d50289e54ed4af110f853360a5f5");

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
                    localAD = uc.RetrieveAuditChecklist(auditList[i].audit_id, authToken);

                    //*********** UPDATE INFORMATION IN DATABASE ***********
                    dc = new dbConnect();
                    dc.OpenMessageConnection();

                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateEveningChecklist";
                    cmd.Parameters.AddWithValue("@bizUnit", "ym");
                    cmd.Parameters.AddWithValue("@dateSent", localAD.auditDate);
                    cmd.Parameters.AddWithValue("@route", localAD.auditRoute.Replace("Route", "RT"));
                    cmd.Parameters.AddWithValue("@auditID", localAD.auditID);

                    cmd.ExecuteNonQuery();

                }

            }

        }



        public void RunDailyRouteWagesMove()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            try
            {
                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                MoveDailyWages mdw = new MoveDailyWages();

                mdw = gc.ConnectSheetsMoveDailyWage();

                string json = "{'text': '";
                foreach (MoveDailyWage m in mdw.DailyWages)
                {
                    json = json + "Date: " + m.RouteDate + "\nRoute: " + m.RouteNumber + "\nCrew Lead: " + m.CrewLead + "\n";


                    double result;

                    result = Convert.ToDouble(m.WagePercentage.Replace("%", ""));
                    if (result / 100 >= .25)
                    {
                        //wage percentage is above 25%
                        json = json + "*Wage %: " + m.WagePercentage + "*\n";
                    }
                    else
                    {
                        json = json + "Wage %: " + m.WagePercentage + "\n";
                    }


                    result = Convert.ToDouble(m.RPH.Replace("$", ""));
                    if (result < 60)
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

                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movewages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_rundailyroutewagesmove today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }


        public void RunDailyMoveDamage()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            try
            {
                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                MoveWagesMessageModel mwm = new MoveWagesMessageModel();

                mwm = gc.ConnectSheetsYouMoveMeDamage();

                string json = "{'text': 'Date Range:  " + mwm.StartDate + " - " + mwm.EndDate + "\nPop-ins Done: " + mwm.PopinsDone + "\nPop-ins Jobs: " + mwm.PopinsJobs + "\nPop-ins: " + mwm.PopinPercentage + " (" + mwm.PopinGoalPercentage + ")\nDamage %: " + mwm.DamagePercentage + " (" + mwm.DamagePercentageGoal + ")\nDamage: " + mwm.Damage + " (" + mwm.DamageGoal + ")\nService %: " + mwm.ServicePercentage + " (" + mwm.ServicePercentageGoal + ")\nService: " + mwm.Service + " (" + mwm.ServiceGoal + ")'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedamage").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_rundailymovedamage today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }

        public void RunDailyWagePercentageMove()
        {
            CommonClient cc = new CommonClient();
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            try
            {
                string startDate = cc.GetStartDate();
                DateTime start = Convert.ToDateTime(startDate);

                GoogleSheetsConnect gc = new GoogleSheetsConnect();
                MoveWagesMessageModel mwm = new MoveWagesMessageModel();

                mwm = gc.ConnectSheets();

                //Google Reviews Received
                Review r = new Review();
                int NumberOfReviews = r.ReturnReviewsNumber("ym", start.Month.ToString(), start.Year.ToString());

                //Review Requests Sent
                double ReviewRequestsSent = r.ReturnReviewsRequested("ymmjd");

                string json = "{'text': 'Date Range:  " + mwm.StartDate + " - " + mwm.EndDate + "\nRevenue: " + String.Format("{0:C}", mwm.Revenue) + "\nRevenue Pace: " + String.Format("{0:C}", mwm.RevenuePace) + "\nNet Wage: " + mwm.NetWagePercentage + " (" + mwm.GoalWagePercentage + ")\nOH Wage: " + mwm.OHWagePercentage + " (" + mwm.OHGoalWagePercentage + ")\nReviews Requested: " + ReviewRequestsSent + "\nReviews Received: " + NumberOfReviews + "\n'}";
                sc.SendMessage(json, cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movewages").ToString().Replace("\\\\", "\\"));
            }
            catch
            {
                sc.SendMessage("{'text': 'There was an error with slack_rundailywagepercentagemove today'}", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_andrewprivate").ToString().Replace("\\\\", "\\"));
            }

        }

        
        public void RunDailyHealthCheckImport()
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();

            //get schedule/employee information
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();
            List<JunkDailySchedule> jds = new List<JunkDailySchedule>();
            jds = gsc.ConnectSheetsMoveDailySchedule();


            //import to daily sheet
            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            reader.ImportDailyHealthCheckSheet("1dFOiu7ZUEs2EpgQxEFmTjO599AFQsw5j5vYqa9rn7Uc", jds, "YMM");



        }


        public void SendSCLeadsAWelcomeText(bool testing)
        {
            TimeSpan startAfternoon = new TimeSpan(0, 0, 0); //Midnight
            TimeSpan endAfternoon = new TimeSpan(23, 59, 59); //11:59:59 pm
            TimeSpan startMorning = new TimeSpan(0, 0, 0); //Midnight
            TimeSpan endMorning = new TimeSpan(23, 59, 59); //11:59:59 pm
            TimeSpan now = DateTime.Now.TimeOfDay;
            bool pleaseContinue = false;

            //only run on this schedule
            //Sunday 3pm - 10am
            //Monday - Friday 6pm - 8am
            //Saturday 5pm - 9am

            DateTime timeTemp = DateTime.Now;
            switch (timeTemp.DayOfWeek.ToString())
            {
                case "Sunday":
                    endMorning = new TimeSpan(9, 59, 59); //9:59:59 am
                    startAfternoon = new TimeSpan(15, 0, 0); //3:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Monday":
                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Tuesday":
                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Wednesday":
                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Thursday":
                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Friday":
                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
                case "Saturday":
                    endMorning = new TimeSpan(8, 59, 59); //8:59:59 am
                    startAfternoon = new TimeSpan(17, 0, 0); //5:00:00 pm

                    pleaseContinue = (((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon))));
                    break;
            }



            //testing ....
            pleaseContinue = true;



            if (pleaseContinue)
            {
                string messageToSend = "";

                SlackClient sc = new SlackClient();
                ConfigSettings cs = new ConfigSettings();
                string rcTokenURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL").ToString().Replace("\\\\", "\\");
                string rcTokenAuth = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth").ToString().Replace("\\\\", "\\");

                //string rcTokenData = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataMove").ToString().Replace("\\\\", "\\");
                //string rcJunkURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMoveURL").ToString().Replace("\\\\", "\\");
                //string rcJunkPhone = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMovePhone").ToString().Replace("\\\\", "\\");

                string rcTokenData = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenData360Wow").ToString().Replace("\\\\", "\\");
                string rcJunkURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMS360WowURL").ToString().Replace("\\\\", "\\");
                string rcJunkPhone = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMS360WowPhone").ToString().Replace("\\\\", "\\");

                string rcImageLocation = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralMMSImageLocation").ToString().Replace("\\\\", "\\").Replace("\\", "\\\\");
                string scDailySchedule = cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedailyschedule").ToString().Replace("\\\\", "\\");

                //get ring central token
                RingCentralClient rc = new RingCentralClient();
                string accessToken = rc.GetRingCentralToken(rcTokenURL, rcTokenData, rcTokenAuth);

                //get schedule/employee information
                GoogleSheetsConnect gsc = new GoogleSheetsConnect();
                List<JunkDailySchedule> jds = new List<JunkDailySchedule>();
                jds = gsc.ConnectSheetsSCAfterHoursLeads();

                //send text message
                string scMessage = "{'text': 'Messages Sent: \n";
                foreach (JunkDailySchedule js in jds)
                {
                    messageToSend = "Hello " + js.EmpName + "!\\n\\nThank you for submitting an online estimate request.Our sales center is currently closed for the night. Feel free to reply to this message with a good time to talk and one of our sales agents will reach out to you tomorrow.\\n\\nThanks!";
                    if (testing)
                    {
                        //Production
                        //rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "16122813895", messageToSend, "just_move_truck.png");
                        rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "17634380871", messageToSend, "just_move_truck.png");
                    }
                    else
                    {
                        //Production
                        rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "1" + js.EmpPhone, messageToSend, "just_move_truck.png");
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
                    cmd.Parameters.AddWithValue("@bizUnit", "ym");
                    cmd.Parameters.AddWithValue("@dateSent", DateTime.Now.ToString());
                    cmd.Parameters.AddWithValue("@phoneTo", "1" + js.EmpPhone);
                    cmd.Parameters.AddWithValue("@truckTeamLead", "c0d913d3-8ae3-43d2-b95d-7d2effb24395");      //Anthony
                    cmd.Parameters.AddWithValue("@mediaURL", "move_truck.png");
                    cmd.Parameters.AddWithValue("@messageType", "scoffhourslead");
                    cmd.Parameters.AddWithValue("@message", messageToSend);
                    cmd.Parameters.AddWithValue("@sentby", "c0d913d3-8ae3-43d2-b95d-7d2effb24395");     //Anthony

                    cmd.ExecuteNonQuery();

                }
                scMessage = scMessage + "'}";

                //send to slack
                sc.SendMessage(scMessage, scDailySchedule);

            }


        }



        public void RunDailySchedule(bool testing)
        {
            SlackClient sc = new SlackClient();
            ConfigSettings cs = new ConfigSettings();
            string rcTokenURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenURL").ToString().Replace("\\\\", "\\");
            string rcTokenData = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenDataMove").ToString().Replace("\\\\", "\\");
            string rcTokenAuth = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralTokenAuth").ToString().Replace("\\\\", "\\");
            string rcJunkURL = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMoveURL").ToString().Replace("\\\\", "\\");
            string rcJunkPhone = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralSMSMovePhone").ToString().Replace("\\\\", "\\");
            string rcImageLocation = cs.ReturnConfigSetting("NightlyRouteToSlack", "ringCentralMMSImageLocation").ToString().Replace("\\\\", "\\").Replace("\\", "\\\\");
            string scDailySchedule = cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movedailyschedule").ToString().Replace("\\\\", "\\");

            //get ring central token
            RingCentralClient rc = new RingCentralClient();
            string accessToken = rc.GetRingCentralToken(rcTokenURL, rcTokenData, rcTokenAuth);

            //get schedule/employee information
            GoogleSheetsConnect gsc = new GoogleSheetsConnect();
            List<JunkDailySchedule> jds = new List<JunkDailySchedule>();
            jds = gsc.ConnectSheetsMoveDailySchedule();

            //send text message
            string scMessage = "{'text': 'Messages Sent: \n";
            foreach (JunkDailySchedule js in jds)
            {
                if (testing)
                {
                    //Production
                    rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "16122813895", "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + ".", "move_truck.png");
                }
                else
                {
                    //Production
                    rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "1" + js.EmpPhone, "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + ".", "move_truck.png");
                }

                //Dev send to myself
                //rc.SendMMSRingCentral(accessToken, rcJunkURL, rcJunkPhone, rcImageLocation, "16122813895", "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + ".", "move_truck.png");

                scMessage = scMessage + "   " + js.EmpName + " - " + js.EmpPhone + "  " + js.EmpStartTime + " " + js.EmpWorkType + "\n";

                //insert into database
                dbConnect dc = new dbConnect();
                dc.OpenMessageConnection();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dc.conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spInsertSent";
                cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                cmd.Parameters.AddWithValue("@bizUnit", "ym");
                cmd.Parameters.AddWithValue("@dateSent", DateTime.Now.ToString());
                cmd.Parameters.AddWithValue("@phoneTo", "1" + js.EmpPhone);
                cmd.Parameters.AddWithValue("@truckTeamLead", "c0d913d3-8ae3-43d2-b95d-7d2effb24395");      //Anthony
                cmd.Parameters.AddWithValue("@mediaURL", "move_truck.png");
                cmd.Parameters.AddWithValue("@messageType", "dailyschedule");
                cmd.Parameters.AddWithValue("@message", "Good evening " + js.EmpName + ", your scheduled start time tomorrow is at " + js.EmpStartTime + ".");
                cmd.Parameters.AddWithValue("@sentby", "c0d913d3-8ae3-43d2-b95d-7d2effb24395");     //Anthony

                cmd.ExecuteNonQuery();


                if (!js.EmpWorkType.ToLower().Contains("team lead") & !js.EmpWorkType.ToLower().Contains("on-call") & !js.EmpWorkType.ToLower().Contains("point") & !string.IsNullOrEmpty(js.EmpWorkType.Trim()))
                {
                    cmd = new SqlCommand();
                    cmd.Connection = dc.conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertRouteChecklists";
                    cmd.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@bizUnit", "ym");
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


    }
}
