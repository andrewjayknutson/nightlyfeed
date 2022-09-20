using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Configuration;
using Newtonsoft.Json;

namespace NightlyRouteToSlack.Utilities
{
    public class GoogleSheetsConnect
    {

        public MoveDailyWages ConnectSheetsShackPopInsTracking(string sheetID)
        {
            int counter = 0;
            int firstLine = 0;
            int endLine = 0;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "Sheet2!A1:F100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //overall NPS score
                mdw.startDate = spreadSheet.Rows[0].ColumnValue(1).ToString() + " - " + spreadSheet.Rows[1].ColumnValue(1).ToString();
                mdw.endDate = spreadSheet.Rows[3].ColumnValue(1).ToString();    //# of jobs
                mdw.goal = spreadSheet.Rows[4].ColumnValue(1).ToString();       //pop in completed
                mdw.avg = spreadSheet.Rows[5].ColumnValue(1).ToString();        //average score
                mdw.score = spreadSheet.Rows[6].ColumnValue(1).ToString();       //percentage

                //find row section of "Lead Tech"
                for (int x = 7; x < 100; x++)
                {
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "lead tech")
                    {
                        firstLine = x;
                    }
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "totals")
                    {
                        endLine = x;
                        break;
                    }
                }


                //rows of Tech Leads
                for (int i = firstLine + 1; i < endLine; i++)
                {
                    if (!(spreadSheet.Rows[i].ColumnValue(0).ToString() == ""))
                    {
                        m = new MoveDailyWage();
                        //Lead Tech
                        m.StartDate = spreadSheet.Rows[i].ColumnValue(0).ToString().Replace("'", "");
                        //Total responses
                        m.EndDate = spreadSheet.Rows[i].ColumnValue(1).ToString();
                        //Average score
                        m.CrewLead = spreadSheet.Rows[i].ColumnValue(3).ToString();
                        mdw.DailyWages.Add(m);
                    }
                }

            }
            catch (WebException ex)
            {

            }

            return mdw;



        }

        public MoveDailyWages ConnectSheetsShackNPSScorecard(string sheetID)
        {
            int counter = 0;
            int firstLine = 0;
            int endLine = 0;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "Sheet1!A1:F100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //overall NPS score
                mdw.startDate = spreadSheet.Rows[0].ColumnValue(1).ToString();
                mdw.endDate = spreadSheet.Rows[1].ColumnValue(1).ToString();
                mdw.score = spreadSheet.Rows[6].ColumnValue(2).ToString();
                mdw.goal = spreadSheet.Rows[6].ColumnValue(3).ToString();

                //find row section of "Lead Tech"
                for(int x = 7; x < 100; x++)
                {
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "lead tech")
                    {
                        firstLine = x;
                    }
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "totals")
                    {
                        endLine = x;
                        break;
                    }
                }


                //rows of Tech Leads
                for (int i = firstLine + 1; i < endLine; i++)
                {
                    if (!(spreadSheet.Rows[i].ColumnValue(0).ToString() == ""))
                    {
                        m = new MoveDailyWage();
                        //Lead Tech
                        m.StartDate = spreadSheet.Rows[i].ColumnValue(0).ToString().Replace("'", "");
                        //NPS score
                        m.EndDate = spreadSheet.Rows[i].ColumnValue(5).ToString();
                        //Total responses
                        m.CrewLead = spreadSheet.Rows[i].ColumnValue(4).ToString();
                        mdw.DailyWages.Add(m);
                    }
                }

                //find row section of "Tech"
                for (int x = endLine + 1; x < 100; x++)
                {
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "tech")
                    {
                        firstLine = x;
                    }
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "totals")
                    {
                        endLine = x;
                        break;
                    }
                }


                m = new MoveDailyWage();
                //Tech
                m.StartDate = "\n   *Techs*\n";
                m.EndDate = "stop";
                mdw.DailyWages.Add(m);


                //rows of Tech
                for (int i = firstLine + 1; i < endLine; i++)
                {
                    if (!(spreadSheet.Rows[i].ColumnValue(0).ToString() == ""))
                    {
                        m = new MoveDailyWage();
                        //Tech
                        m.StartDate = spreadSheet.Rows[i].ColumnValue(0).ToString().Replace("'","");
                        //NPS score
                        m.EndDate = spreadSheet.Rows[i].ColumnValue(5).ToString();
                        //Total responses
                        m.CrewLead = spreadSheet.Rows[i].ColumnValue(4).ToString();
                        mdw.DailyWages.Add(m);
                    }
                }

            }
            catch (WebException ex)
            {

            }

            return mdw;


        }

        public float ConnectSheetsJunkDumpData(string sheetID)
        {
            int rowToCheck = 15;
            float returnVal = 0;

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                //string range = "Daily Reporting!A1:C50";
                string range = "Disposal Reporting By Location!A1:I50";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //Dumping Fees  -> cell I16
                //should be I16 but let's check column H for "dump"
                for (int x = 10; x < 30; x++)
                {
                    if (spreadSheet.Rows[x].ColumnValue(7).ToString().Trim().ToLower().Contains("dump"))
                    {
                        rowToCheck = x;
                        break;
                    }
                }


                if (!(float.TryParse(spreadSheet.Rows[rowToCheck].ColumnValue(8).ToString().Replace("%", ""), out returnVal)))
                {
                    returnVal = -1;
                }
                else
                {
                    returnVal = returnVal / 100;
                }
                //returnVal = float.Parse(spreadSheet.Rows[15].ColumnValue(8).ToString().Replace("%",""));

            }
            catch (WebException ex)
            {

            }

            return returnVal;

        }

        public JunkGoals ConnectSheetsJunkGoals(string sheetID, string firstOfMonthDate)
        {
            JunkGoals m = new JunkGoals();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "gj_goals!A1:E100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                for (int x = 0; x < 100; x++)
                {
                    //grab individual goals
                    switch (spreadSheet.Rows[x].ColumnValue(3).ToString().Trim().ToLower())
                    {
                        case "dailyajs":
                            m.dailyAJS = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailyrph":
                            m.dailyRPH = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailytruckrph":
                            m.dailyTruckRPH = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailysupport%":
                            m.dailySupport = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailydirect%":
                            m.dailyDirectWage = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailysga%":
                            m.dailySGA = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "dailydump%":
                            m.dailyDump = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlyajs":
                            m.monthlyAJS = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlyrph":
                            m.monthlyRPH = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlytruckrph":
                            m.monthlyTruckRPH = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlysupport%":
                            m.monthlySupport = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlydirect%":
                            m.monthlyDirectWage = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlysga%":
                            m.monthlySGA = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                        case "monthlydump%":
                            m.monthlyDump = spreadSheet.Rows[x].ColumnValue(4).ToString();
                            break;
                    }


                    //grab revenue goal and jump out
                    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == firstOfMonthDate)
                    {
                        m.revenue = spreadSheet.Rows[x].ColumnValue(1).ToString();
                        break;
                    }




                }


            }
            catch (WebException ex)
            {

            }

            return m;

        }


        public MoveDailyWage ConnectSheetsJunkSGAData(string sheetID)
        {
            MoveDailyWage m = new MoveDailyWage();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "GJ SG&A Slack Output!A1:C25";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                for (int x = 0; x < spreadSheet.Rows.Count; x++)
                {
                    switch (spreadSheet.Rows[x].ColumnValue(2).ToString().Trim().ToLower())
                    {
                        case "sgaperday":
                            //SG&A $'s Per Day
                            m.StartDate = spreadSheet.Rows[x].ColumnValue(1).ToString();
                            break;
                        case "totalused":
                            //Total Used
                            m.EndDate = spreadSheet.Rows[x].ColumnValue(1).ToString();
                            break;
                        case "sgagoal":
                            //SGA % Goal
                            m.MiscItem = spreadSheet.Rows[x].ColumnValue(1).ToString();
                            break;
                        case "directwagegoal":
                            //Direct Wage % Goal
                            m.MiscItem2 = spreadSheet.Rows[x].ColumnValue(1).ToString();
                            break;
                    }

                }


            }
            catch (WebException ex)
            {

            }

            return m;

        }


        public MoveDailyWages ConnectSheetsJunkNPSScorecard(string sheetID)
        {
            int counter = 0;
            int firstLine = 0;
            int endLine = 0;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                //string range = "NPS Reporting!A1:F100";
                //string range = "slack_reporting!A1:K200";
                string range = "NPS Reporting - TC!A1:K200";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //overall NPS score
                mdw.startDate = spreadSheet.Rows[0].ColumnValue(1).ToString();
                mdw.endDate = spreadSheet.Rows[1].ColumnValue(1).ToString();
                //mdw.thruDate = spreadSheet.Rows[0].ColumnValue(3).ToString();
                //mdw.lastUpdated = spreadSheet.Rows[1].ColumnValue(3).ToString();
                //mdw.score = spreadSheet.Rows[6].ColumnValue(2).ToString();
                //mdw.goal = spreadSheet.Rows[6].ColumnValue(3).ToString();

                //find row section of "Team Member"
                //for (int x = 7; x < 100; x++)
                //{
                //    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "team member")
                //    {
                //        firstLine = x;
                //    }
                //    if (spreadSheet.Rows[x].ColumnValue(0).ToString().Trim().ToLower() == "totals")
                //    {
                //        endLine = x;
                //        break;
                //    }
                //}


                //firstLine = 5;
                firstLine = 11;
                endLine = 200;

                //rows of Team Members
                for (int i = firstLine; i < spreadSheet.Rows.Count; i++)
                {
                    if (!(spreadSheet.Rows[i].ColumnValue(0).ToString().Trim() == "") & !(spreadSheet.Rows[i].ColumnValue(0).ToString().Trim().ToLower() == "totals"))
                    {
                        m = new MoveDailyWage();
                        //Team Member
                        m.StartDate = spreadSheet.Rows[i].ColumnValue(0).ToString();
                        //NPS score
                        m.EndDate = spreadSheet.Rows[i].ColumnValue(5).ToString();
                        //Total responses
                        m.CrewLead = spreadSheet.Rows[i].ColumnValue(4).ToString();
                        ////Team Member
                        //m.StartDate = spreadSheet.Rows[i].ColumnValue(0).ToString();
                        ////NPS score
                        //m.EndDate = spreadSheet.Rows[i].ColumnValue(5).ToString();
                        ////Total responses
                        //m.CrewLead = spreadSheet.Rows[i].ColumnValue(4).ToString();
                        mdw.DailyWages.Add(m);
                    }
                    else
                    {
                        break;
                    }

                }

            }
            catch (WebException ex)
            {

            }

            return mdw;


        }




        public bool DidJakeWorkTodayWithoutTimeCardEntry()
        {
            bool startLooking = false;
            bool isWorking = false;
            bool hoursEntered = false;
            string todayDay = DateTime.Now.DayOfWeek.ToString();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);
                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);


                //**** Is he on the schedule for today?
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader("1MTwqtKym02eP6xHX3OmmbnfNtxG9CocS_YCTrQ_rqiM", "Schedule!A1:X50");

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        switch (row.ColumnValue(0).ToString().Trim().ToLower())
                        {
                            case "monday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break;

                            case "tuesday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break; 

                            case "wednesday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break; 

                            case "thursday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break; 

                            case "friday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break; 

                            case "saturday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break;

                            case "sunday":
                                startLooking = (todayDay.Trim().ToLower() == row.ColumnValue(0).ToString().Trim().ToLower());
                                break;

                            default:
                                break;
                        }

                        if(startLooking & row.ColumnValue(0).ToString().Trim().ToLower() == "jake")
                        {
                            //I've found the row in which Jake is specified as working or not today
                            for(int i = 1; i < 22; i++)
                            {
                                isWorking = (row.ColumnValue(i).ToString().Trim().ToLower() == "true");
                                if (isWorking) { break; }
                            }

                        }

                    }
                    catch (WebException ex)
                    {

                    }
                }




                //Has he entered any hours yet today?
                spreadSheet = reader.GetSpreadSheetNoHeader("1PLl4LBAZ2ov0ebt4RSnv0Vh39_QGQV01fPZTb1Vr5ss", "HOURS SUMMARY!A1:E400");

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    if (row.ColumnValue(0).ToString().Trim() == DateTime.Now.ToShortDateString())
                    {
                        hoursEntered = true;
                        break;
                    }
                }

            }
            catch (WebException ex)
            {

            }

            //return value of whether or not he worked and didn't enter his hours
            return (isWorking & !hoursEntered);
        }


        public MoveDailyWages ConnectSheetsMoveDamage(string sheetID)
        {
            string startDate = "";
            string endDate = "";
            bool startGrabbingData = false;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "Reporting!S1:V200";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        //set startDate
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "start date:")
                        {
                            startDate = row.ColumnValue(1).ToString();
                        }

                        //set endDate
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "finish date:")
                        {
                            endDate = row.ColumnValue(1).ToString();
                        }

                        if (startGrabbingData & !(row.ColumnValue(0).ToString().Trim() ==""))
                        {
                            m.StartDate = startDate;
                            m.EndDate = endDate;
                            m.CrewLead = row.ColumnValue(0).ToString().Trim();

                            //Damage
                            m.MiscItem = row.ColumnValue(1).ToString().Trim();

                            //Service
                            m.MiscItem2 = row.ColumnValue(2).ToString().Trim();

                            //Total
                            m.MiscItem3 = row.ColumnValue(3).ToString().Trim();

                            mdw.DailyWages.Add(m);
                            m = new MoveDailyWage();
                        }

                        //set startGrabbingData
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "crew leader")
                        {
                            startGrabbingData = true;
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



            return mdw;

        }



        public MoveDailyWages ConnectSheetsSalesCenterConversion(string sheetID)
        {
            string startDate = "";
            string endDate = "";
            string estimator = "";
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            //Get the Data in the current sheet
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "TD_Conversion!A1:L200";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        //set startDate and endDate
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "set custom range (7 day)")
                        {
                            startDate = row.ColumnValue(1).ToString();
                            endDate = row.ColumnValue(2).ToString();
                        }

                        //set Estimator value
                        switch (row.ColumnValue(0).ToString().Trim().ToLower())
                        {
                            case "combined conversion":
                                estimator = "Combined Conversion";
                                break;
                            case "jacob":
                                estimator = "Jacob";
                                break;
                            case "jake":
                                estimator = "Jake";
                                break;
                            case "national sc":
                                estimator = "National SC";
                                break;
                            case "james":
                                estimator = "James";
                                break;
                            case "haylee":
                                estimator = "Haylee";
                                break;
                            default:
                                break;
                        }

                        //startDate -> StartDate
                        //endDate -> EndDate
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


                        //set when find 30 day
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "30 day")
                        {
                            m.StartDate = startDate;
                            m.EndDate = endDate;
                            m.MiscItem = estimator;

                            m.CrewLead = row.ColumnValue(3).ToString().Trim();
                            m.WagePercentage = row.ColumnValue(4).ToString().Trim();
                            m.RPH = row.ColumnValue(5).ToString().Trim();
                            m.RouteDate = row.ColumnValue(6).ToString().Trim();
                            m.RouteNumber = row.ColumnValue(7).ToString().Trim();

                        }

                        //set when find date range
                        if (row.ColumnValue(0).ToString().Trim().ToLower() == "custom range (7 day)")
                        {
                            m.StartDate = startDate;
                            m.EndDate = endDate;
                            m.MiscItem = estimator;

                            m.MiscItem2 = row.ColumnValue(3).ToString().Trim();
                            m.MiscItem3 = row.ColumnValue(4).ToString().Trim();
                            m.MiscItem4 = row.ColumnValue(5).ToString().Trim();
                            m.MiscItem5 = row.ColumnValue(6).ToString().Trim();
                            m.MiscItems6 = row.ColumnValue(7).ToString().Trim();

                            mdw.DailyWages.Add(m);
                            m = new MoveDailyWage();
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

            return mdw;


        }



        public MoveDailyWage ConnectSheetsShackDailyScorecard(DateTime startDate, DateTime endDate)
        {
            int PmtRcd = 24;
            int RPHBonus = 66;
            int TotalSales = 69;
            int ProdHours = 71;
            int Wages = 72;
            int firstTotalSales = 16;
            int oppID = 8;

            int dataIndex = 0;
            DateTime temp;
            bool startCounting = false;
            int totalAJS = 0;

            string totalRevenue = "0";
            string totalWages = "0";
            string totalMovingHours = "0";
            string totalPaymentReceived = "0";
            string totalRPHBonus = "0";

            MoveDailyWage m = new MoveDailyWage();
            DateTime now = DateTime.Now;

            //the intent is to cycle through the sheet and get the total revenue for the day before 


            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1ttCMCGo-C5ozvlgWgMvQKkSY0Ujw4hO_yJlL3qLOn_Q";          //Shack Payroll & Labor
                string range = "Job Tracking!A1:BU8000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        DateTime.TryParse(row.ColumnValue(1).ToString().Trim(), out temp);
                        if (temp >= startDate && temp <= endDate)
                        {
                            startCounting = true;
                        }
                        if (startCounting)
                        {
                            //Pmt R'cvd
                            if (row.ColumnValue(PmtRcd).ToString().IndexOf("-") == -1)
                            {
                                if (!(row.ColumnValue(PmtRcd).ToString().Trim().Replace("$", "").Replace(" ", "").Replace("-", "") == ""))
                                {
                                    totalPaymentReceived = (float.Parse(totalPaymentReceived) + float.Parse(row.ColumnValue(PmtRcd).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }
                            }

                            //RPH Bonus (to be added to total wages)
                            if (row.ColumnValue(RPHBonus).ToString().IndexOf("-") == -1)
                            {
                                if (!(row.ColumnValue(RPHBonus).ToString().Trim().Replace("$", "").Replace(" ", "").Replace("-", "") == ""))
                                {
                                    totalRPHBonus = (float.Parse(totalRPHBonus) + float.Parse(row.ColumnValue(RPHBonus).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }
                            }


                            //keep track of how many jobs are done
                            if (row.ColumnValue(firstTotalSales).ToString().IndexOf("-") == -1 && !(row.ColumnValue(oppID).ToString().ToLower().Trim() == "totals"))
                            {
                                totalAJS = totalAJS + 1;
                            }



                            if (dataIndex == 0)
                            {
                                //Total Sales
                                if (row.ColumnValue(TotalSales).ToString().IndexOf("-") == -1)
                                {
                                    totalRevenue = (float.Parse(totalRevenue) + float.Parse(row.ColumnValue(TotalSales).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }

                                //Production hours
                                if (row.ColumnValue(ProdHours).ToString().IndexOf("-") == -1)
                                {
                                    totalMovingHours = (float.Parse(totalMovingHours) + float.Parse(row.ColumnValue(ProdHours).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }

                                dataIndex = dataIndex + 1;
                            }
                            else if (dataIndex == 5)
                            {

                                //Total Wages     
                                if (row.ColumnValue(Wages).ToString().IndexOf("-") == -1)
                                {
                                    totalWages = (float.Parse(totalWages) + float.Parse(row.ColumnValue(Wages).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }

                                dataIndex = dataIndex + 1;
                            }
                            else if (dataIndex == 6)
                            {
                                startCounting = false;
                                dataIndex = 0;

                            }
                            else
                            {
                                dataIndex = dataIndex + 1;
                            }
                        }

                    }
                    catch (WebException ex)
                    {

                    }



                }


                range = "Wage Management!A1:K100";

                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Wage%
                m.MiscItem = spreadSheet.Rows[5].ColumnValue(1).ToString().Trim();

                //Wage% goal
                m.MiscItem2 = spreadSheet.Rows[6].ColumnValue(1).ToString().Trim();

                //OH%
                m.MiscItem3 = spreadSheet.Rows[16].ColumnValue(10).ToString().Trim();

                //OH% goal
                m.MiscItem4 = spreadSheet.Rows[2].ColumnValue(10).ToString().Trim();



            }
            catch (WebException ex)
            {

            }

            //Payment Received
            m.CrewLead = totalPaymentReceived.ToString();

            //Revenue
            m.RouteDate = totalRevenue.ToString();

            //Production Wage %
            totalWages = (float.Parse(totalWages) + float.Parse(totalRPHBonus)).ToString();
            m.RouteNumber = (float.Parse(totalWages.Replace("$", "").Replace(" ", "")) / float.Parse(totalRevenue.Replace("$", "").Replace(" ", ""))).ToString();

            //RPH
            m.RPH = (float.Parse(totalRevenue.Replace("$", "").Replace(" ", "")) / float.Parse(totalMovingHours.Replace("$", "").Replace(" ", ""))).ToString();

            //AJS
            m.WagePercentage = (float.Parse(totalRevenue.Replace("$", "").Replace(" ", "")) / totalAJS).ToString(); ;

            return m;

        }











        public MoveDailyWage ConnectSheetsGJDailyRecon(string sheetID)
        {
            DateTime Temp;
            MoveDailyWage mdw = new MoveDailyWage();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = sheetID;
                string range = "2020_recon!A1:I500";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(spreadSheetId, range);

                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    if (DateTime.TryParse(row.ColumnValue(2).ToString(), out Temp) == true) 
                    {
                        if (DateTime.Parse(row.ColumnValue(2).ToString()).ToShortDateString() == DateTime.Now.AddDays(-1).ToShortDateString())
                        {
                            mdw = new MoveDailyWage();

                            //Closer
                            mdw.CrewLead = row.ColumnValue(1).ToString();
                            //Route Date
                            mdw.WagePercentage = row.ColumnValue(2).ToString();
                            //Pipeline
                            mdw.RouteNumber = row.ColumnValue(3).ToString();
                            //Route Management
                            mdw.RouteDate = row.ColumnValue(4).ToString();
                            //Pipe/RM Difference
                            mdw.RPH = row.ColumnValue(6).ToString();

                            break;
                        }

                    }



                }
            }
            catch (WebException ex)
            {

            }



            return mdw;
        }



        public MoveDailyWages ConnectSheetsWeeklyScorecard(string sheetID, string bizTab)
        {
            MoveDailyWage mdw = new MoveDailyWage();
            MoveDailyWages localMDW = new MoveDailyWages();
            localMDW.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = sheetID;
                string range = bizTab + "!A1:D50";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(spreadSheetId, range);

                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {

                    if ((row.ColumnValue(0).ToString() == ""))
                    { break; }
                    else
                    {
                        mdw = new MoveDailyWage();

                        //responsible
                        mdw.CrewLead = row.ColumnValue(0).ToString();
                        //item
                        mdw.RPH = row.ColumnValue(1).ToString();
                        //goal
                        mdw.RouteDate = row.ColumnValue(2).ToString();
                        //actual
                        mdw.RouteNumber = row.ColumnValue(3).ToString();

                        localMDW.DailyWages.Add(mdw);
                    }
                }
            }
            catch (WebException ex)
            {

            }


            return localMDW;

        }






        public MoveDailyWages ConnectSheetsVISATransactions(string sheetID, string teamMember, DateTime startDate, DateTime endDate)
        {
            MoveDailyWage mdw = new MoveDailyWage();
            MoveDailyWages localMDW = new MoveDailyWages();
            localMDW.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = sheetID;
                string range = "Transactions!A1:D5000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    if (row.ColumnValue(2).ToString().Trim().ToLower() == teamMember.Trim().ToLower())
                    {
                        if (Convert.ToDateTime(row.ColumnValue(0).ToString()) >= startDate && Convert.ToDateTime(row.ColumnValue(0).ToString()) <= endDate)
                        {
                            mdw = new MoveDailyWage();

                            //Date
                            mdw.RouteDate = row.ColumnValue(0).ToString();
                            //Amount
                            mdw.RouteNumber = row.ColumnValue(1).ToString();
                            //Vendor
                            mdw.CrewLead = row.ColumnValue(3).ToString();

                            localMDW.DailyWages.Add(mdw);

                        }
                    }
                }


                ////get Start Date
                //SpreadSheetRow row = spreadSheet.Rows[24];
                //mwm.StartDate = row.ColumnValue(0).ToString();

                ////get End Date
                //row = spreadSheet.Rows[25];
                //mwm.EndDate = row.ColumnValue(0).ToString();

                ////get Revenue
                //row = spreadSheet.Rows[62];
                //mwm.Revenue = row.ColumnValue(8).ToString();

                ////get RevenuePace = Gross Overhead
                //row = spreadSheet.Rows[41];
                //mwm.RevenuePace = row.ColumnValue(8).ToString();

                ////get Gross Overhead Percentage
                //row = spreadSheet.Rows[42];
                //mwm.GrossOverheadPercentage = row.ColumnValue(8).ToString();

                ////get NetWagePercentage = Move Hours Cost Reduction
                //row = spreadSheet.Rows[44];
                //mwm.NetWagePercentage = row.ColumnValue(8).ToString();

                ////get GoalWagePercentage = Junk/Shack Cost Reduction
                //row = spreadSheet.Rows[55];
                //mwm.GoalWagePercentage = row.ColumnValue(8).ToString();

                ////get OHWagePercentage = Net Overhead
                //row = spreadSheet.Rows[58];
                //mwm.OHWagePercentage = row.ColumnValue(8).ToString();

                ////get Net Overhead Percentage
                //row = spreadSheet.Rows[58];
                //mwm.NetOverheadPercentage = row.ColumnValue(8).ToString();

                ////get Overhead Goal
                //row = spreadSheet.Rows[60];
                //mwm.OverheadGoal = row.ColumnValue(8).ToString();

                ////get Overhead Goal Results
                //row = spreadSheet.Rows[67];
                //mwm.OverheadGoalResults = row.ColumnValue(8).ToString();

                ////get OHGoalWagePercentage = Cash Flow Impact
                //row = spreadSheet.Rows[66];
                //mwm.OHGoalWagePercentage = row.ColumnValue(8).ToString();

            }
            catch (WebException ex)
            {

            }


            return localMDW;

        }


        public MoveDailyWages ConnectSheetsVISACardholders(string sheetID)
        {
            MoveDailyWage mdw = new MoveDailyWage();
            MoveDailyWages localMDW = new MoveDailyWages();
            localMDW.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = sheetID;
                string range = "Cardholders!A1:A50";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    mdw = new MoveDailyWage();
                    mdw.CrewLead = row.ColumnValue(0).ToString();

                    localMDW.DailyWages.Add(mdw);
                }


                ////get Start Date
                //SpreadSheetRow row = spreadSheet.Rows[24];
                //mwm.StartDate = row.ColumnValue(0).ToString();

                ////get End Date
                //row = spreadSheet.Rows[25];
                //mwm.EndDate = row.ColumnValue(0).ToString();

                ////get Revenue
                //row = spreadSheet.Rows[62];
                //mwm.Revenue = row.ColumnValue(8).ToString();

                ////get RevenuePace = Gross Overhead
                //row = spreadSheet.Rows[41];
                //mwm.RevenuePace = row.ColumnValue(8).ToString();

                ////get Gross Overhead Percentage
                //row = spreadSheet.Rows[42];
                //mwm.GrossOverheadPercentage = row.ColumnValue(8).ToString();

                ////get NetWagePercentage = Move Hours Cost Reduction
                //row = spreadSheet.Rows[44];
                //mwm.NetWagePercentage = row.ColumnValue(8).ToString();

                ////get GoalWagePercentage = Junk/Shack Cost Reduction
                //row = spreadSheet.Rows[55];
                //mwm.GoalWagePercentage = row.ColumnValue(8).ToString();

                ////get OHWagePercentage = Net Overhead
                //row = spreadSheet.Rows[58];
                //mwm.OHWagePercentage = row.ColumnValue(8).ToString();

                ////get Net Overhead Percentage
                //row = spreadSheet.Rows[58];
                //mwm.NetOverheadPercentage = row.ColumnValue(8).ToString();

                ////get Overhead Goal
                //row = spreadSheet.Rows[60];
                //mwm.OverheadGoal = row.ColumnValue(8).ToString();

                ////get Overhead Goal Results
                //row = spreadSheet.Rows[67];
                //mwm.OverheadGoalResults = row.ColumnValue(8).ToString();

                ////get OHGoalWagePercentage = Cash Flow Impact
                //row = spreadSheet.Rows[66];
                //mwm.OHGoalWagePercentage = row.ColumnValue(8).ToString();

            }
            catch (WebException ex)
            {

            }


            return localMDW;

        }


        public MoveWagesMessageModel ConnectSheetsOverheadWages()
        {
            MoveWagesMessageModel mwm = new MoveWagesMessageModel();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I";
                string range = "Wage Management!A1:J100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //get Start Date
                SpreadSheetRow row = spreadSheet.Rows[24];
                mwm.StartDate = row.ColumnValue(0).ToString();

                //get End Date
                row = spreadSheet.Rows[25];
                mwm.EndDate = row.ColumnValue(0).ToString();

                //get Revenue
                row = spreadSheet.Rows[62];
                mwm.Revenue = row.ColumnValue(8).ToString();

                //get RevenuePace = Gross Overhead
                row = spreadSheet.Rows[41];
                mwm.RevenuePace = row.ColumnValue(8).ToString();

                //get Gross Overhead Percentage
                row = spreadSheet.Rows[42];
                mwm.GrossOverheadPercentage = row.ColumnValue(8).ToString();

                //get NetWagePercentage = Move Hours Cost Reduction
                row = spreadSheet.Rows[44];
                mwm.NetWagePercentage = row.ColumnValue(8).ToString();

                //get GoalWagePercentage = Junk/Shack Cost Reduction
                row = spreadSheet.Rows[55];
                mwm.GoalWagePercentage = row.ColumnValue(8).ToString();

                //get OHWagePercentage = Net Overhead
                row = spreadSheet.Rows[58];
                mwm.OHWagePercentage = row.ColumnValue(8).ToString();

                //get Net Overhead Percentage
                row = spreadSheet.Rows[58];
                mwm.NetOverheadPercentage = row.ColumnValue(8).ToString();

                //get Overhead Goal
                row = spreadSheet.Rows[60];
                mwm.OverheadGoal = row.ColumnValue(8).ToString();

                //get Overhead Goal Results
                row = spreadSheet.Rows[67];
                mwm.OverheadGoalResults = row.ColumnValue(8).ToString();

                //get OHGoalWagePercentage = Cash Flow Impact
                row = spreadSheet.Rows[66];
                mwm.OHGoalWagePercentage = row.ColumnValue(8).ToString();

            }
            catch (WebException ex)
            {

            }


            return mwm;

        }



        public MoveWagesMessageModel ConnectSheets()
        {
            MoveWagesMessageModel mwm = new MoveWagesMessageModel();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I";
                string range = "Wage Management!A1:J100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //get Start Date
                SpreadSheetRow row = spreadSheet.Rows[24];
                mwm.StartDate = row.ColumnValue(0).ToString();

                //get End Date
                row = spreadSheet.Rows[25];
                mwm.EndDate = row.ColumnValue(0).ToString();

                //get Revenue
                row = spreadSheet.Rows[4];
                mwm.Revenue = row.ColumnValue(1).ToString();

                //get Revenue Pace
                row = spreadSheet.Rows[14];
                mwm.RevenuePace = row.ColumnValue(1).ToString();

                //get Net Wage %
                row = spreadSheet.Rows[5];
                mwm.NetWagePercentage = row.ColumnValue(1).ToString();

                //get Wage % Goal
                row = spreadSheet.Rows[6];
                mwm.GoalWagePercentage = row.ColumnValue(1).ToString();

                //get OH Wage %
                row = spreadSheet.Rows[7];
                mwm.OHWagePercentage = row.ColumnValue(6).ToString();

                //get OH Wage % Goal
                row = spreadSheet.Rows[2];
                mwm.OHGoalWagePercentage = row.ColumnValue(6).ToString();

            }
            catch (WebException ex)
            {

            }


            return mwm;

        }



        public MoveWagesMessageModel ConnectSheetsYouMoveMeDamage()
        {
            MoveWagesMessageModel mwm = new MoveWagesMessageModel();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1Vcavfwan9FbZ0O2En4G6FkQ_6lL_3J1XhF1rtVlgwV0";      //{updated}YMM Customer Follow Up
                //string spreadSheetId = "1V-RV8haKT5tLnwSxzM3INo2oa02VZN0K2QRKDaNoZtc";      //YMM Customer Follow Up
                string range = "Pop Ins!A1:J50";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //get Pop Ins Completed, Pop Ins Jobs and Percentage 
                SpreadSheetRow row = spreadSheet.Rows[0];
                mwm.PopinsDone = row.ColumnValue(2).ToString();
                mwm.PopinsJobs = row.ColumnValue(3).ToString();
                mwm.PopinPercentage = row.ColumnValue(4).ToString();
                mwm.StartDate = row.ColumnValue(1).ToString();

                //get Pop In Goal Percentage
                row = spreadSheet.Rows[1];
                mwm.PopinGoalPercentage = row.ColumnValue(4).ToString();
                mwm.EndDate = row.ColumnValue(1).ToString();


                range = "Reporting!A1:AC50";

                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //get Damage & Service Data
                row = spreadSheet.Rows[4];
                mwm.DamagePercentageGoal = row.ColumnValue(24).ToString();
                mwm.DamagePercentage = row.ColumnValue(25).ToString();

                row = spreadSheet.Rows[5];
                mwm.ServicePercentageGoal = row.ColumnValue(24).ToString();
                mwm.ServicePercentage = row.ColumnValue(25).ToString();

                row = spreadSheet.Rows[8];
                mwm.DamageGoal = row.ColumnValue(24).ToString();
                mwm.Damage = row.ColumnValue(25).ToString();

                row = spreadSheet.Rows[9];
                mwm.ServiceGoal = row.ColumnValue(24).ToString();
                mwm.Service = row.ColumnValue(25).ToString();

            }
            catch (WebException ex)
            {

            }

            return mwm;

        }


        public ShackSales ConnectSheetsShackSales()
        {
            ShackSales mwm = new ShackSales();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1FzbOCWtMoS0YgTa7PSKnrC2WrI7AHMl8jHPm2n8qdJA";
                string range = "Summary by Month!A16:I24";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Detailing 30 Day
                SpreadSheetRow row = spreadSheet.Rows[1];
                mwm.StartDate30Day = row.ColumnValue(1).ToString();
                mwm.EndDate30Day = row.ColumnValue(2).ToString();
                mwm.DetailingEstimates30Day = row.ColumnValue(3).ToString();
                mwm.DetailingConverted30Day = row.ColumnValue(7).ToString();
                mwm.DetailingPercentage30Day = row.ColumnValue(8).ToString();

                //Detailing 7 Day
                row = spreadSheet.Rows[2];
                mwm.StartDate7Day = row.ColumnValue(1).ToString();
                mwm.EndDate7Day = row.ColumnValue(2).ToString();
                mwm.DetailingEstimates7Day = row.ColumnValue(3).ToString();
                mwm.DetailingConverted7Day = row.ColumnValue(7).ToString();
                mwm.DetailingPercentage7Day = row.ColumnValue(8).ToString();

                //Lights 30 Day
                row = spreadSheet.Rows[6];
                mwm.StartDate30Day = row.ColumnValue(1).ToString();
                mwm.EndDate30Day = row.ColumnValue(2).ToString();
                mwm.LightsEstimates30Day = row.ColumnValue(3).ToString();
                mwm.LightsConverted30Day = row.ColumnValue(7).ToString();
                mwm.LightsPercentage30Day = row.ColumnValue(8).ToString();

                //Lights 7 Day
                row = spreadSheet.Rows[7];
                mwm.StartDate7Day = row.ColumnValue(1).ToString();
                mwm.EndDate7Day = row.ColumnValue(2).ToString();
                mwm.LightsEstimates7Day = row.ColumnValue(3).ToString();
                mwm.LightsConverted7Day = row.ColumnValue(7).ToString();
                mwm.LightsPercentage7Day = row.ColumnValue(8).ToString();

            }
            catch (WebException ex)
            {

            }

            return mwm;

        }



        public ShackWagesMessageModel ConnectSheetsShackProductionPace()
        {
            ShackWagesMessageModel mwm = new ShackWagesMessageModel();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1yHwKOLzvI0iv7YGqjuBeO7LIUN9RhC4sLD868g0H0yA";
                string range = "TC Forecast!A1:B100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                ////get Overall Production Pace
                SpreadSheetRow row = spreadSheet.Rows[15];
                mwm.ProductionPace = row.ColumnValue(1).ToString();

                ////get Overall Revenue Goal
                row = spreadSheet.Rows[16];
                mwm.OverallRevenueMTD = row.ColumnValue(1).ToString();

                ////get Lights Production Pace
                row = spreadSheet.Rows[37];
                mwm.LightsRevenuePace = row.ColumnValue(1).ToString();

                ////get Lights Revenue Goal
                row = spreadSheet.Rows[38];
                mwm.LightsRevenue = row.ColumnValue(1).ToString();

                ////get Detailing Production Pace
                row = spreadSheet.Rows[60];
                mwm.DetailingRevenuePace = row.ColumnValue(1).ToString();

                ////get Detailing Revenue Goal
                row = spreadSheet.Rows[61];
                mwm.DetailingRevenue = row.ColumnValue(1).ToString();


            }
            catch (WebException ex)
            {

            }

            return mwm;

        }





        public ShackWagesMessageModel ConnectSheetsShack()
        {
            int currentMonth = 1;
            ShackWagesMessageModel mwm = new ShackWagesMessageModel();

            DateTime now = DateTime.Now;
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

                string spreadSheetId = "1ttCMCGo-C5ozvlgWgMvQKkSY0Ujw4hO_yJlL3qLOn_Q";
                string range = "Wage Management!A1:N100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                float adminOH = 0;
                float OH = 0;
                float mktg = 0;
                float sales = 0;

                //get Start Date
                SpreadSheetRow row = spreadSheet.Rows[23];
                mwm.StartDate = row.ColumnValue(0).ToString();

                //get End Date
                row = spreadSheet.Rows[24];
                mwm.EndDate = row.ColumnValue(0).ToString();

                //get Revenue
                row = spreadSheet.Rows[4];
                mwm.Revenue = row.ColumnValue(1).ToString();

                //get Sales Pace
                row = spreadSheet.Rows[15];
                mwm.SalesPace = row.ColumnValue(7).ToString();

                //get Production Pace
                row = spreadSheet.Rows[16];
                mwm.ProductionPace = row.ColumnValue(7).ToString();

                //get Production Pace
                row = spreadSheet.Rows[14];
                mwm.OverallRevenueMTD = row.ColumnValue(7).ToString();

                //get Net Wage %
                row = spreadSheet.Rows[5];
                mwm.NetWagePercentage = row.ColumnValue(1).ToString();
                mwm.MktgWagePercentage = row.ColumnValue(4).ToString();
                mwm.SalesWagePercentage = row.ColumnValue(7).ToString();

                mktg = float.Parse(row.ColumnValue(4).ToString().Replace("%", ""));
                sales = float.Parse(row.ColumnValue(7).ToString().Replace("%", ""));

                //get Wage % Goal
                row = spreadSheet.Rows[6];
                mwm.GoalWagePercentage = row.ColumnValue(1).ToString();
                mwm.MktgGoalWagePercentage = row.ColumnValue(4).ToString();
                mwm.SalesGoalWagePercentage = row.ColumnValue(7).ToString();

                //get OH Wage %
                row = spreadSheet.Rows[16];
                OH = float.Parse(row.ColumnValue(10).ToString().Replace("%",""));
                adminOH = float.Parse(row.ColumnValue(13).ToString().Replace("%", ""));
                mwm.OHWagePercentage = OH + "%";
                mwm.AdminWagePercentage = adminOH + "%";

                mwm.OverallOHGoalPercentage = "15%";
                mwm.OverallOHPercentage = (mktg + sales + OH + adminOH) + "%";

                //get OH Wage % Goal
                row = spreadSheet.Rows[2];
                OH = float.Parse(row.ColumnValue(10).ToString().Replace("%", ""));
                adminOH = float.Parse(row.ColumnValue(13).ToString().Replace("%", ""));
                mwm.OHGoalWagePercentage = OH + "%";
                mwm.AdminGoalWagePercentage = adminOH + "%";

                //get Lights Revenue
                row = spreadSheet.Rows[24];
                mwm.LightsRevenue = row.ColumnValue(7).ToString();

                //get Lights Revenue Pace
                row = spreadSheet.Rows[25];
                mwm.LightsRevenuePace = row.ColumnValue(7).ToString();

                //get Lights Sales
                row = spreadSheet.Rows[27];
                mwm.LightsSales = row.ColumnValue(7).ToString();

                //get Lights Sales Pace
                row = spreadSheet.Rows[28];
                mwm.LightsSalesPace = row.ColumnValue(7).ToString();

                //get Detailing Sales
                row = spreadSheet.Rows[30];
                mwm.DetailingSales = row.ColumnValue(7).ToString();

                //get Detailing Sales Pace
                row = spreadSheet.Rows[31];
                mwm.DetailingSalesPace = row.ColumnValue(7).ToString();

                //get Detailing Revenue
                row = spreadSheet.Rows[33];
                mwm.DetailingRevenue = row.ColumnValue(7).ToString();

                //get Detailing Revenue Pace
                row = spreadSheet.Rows[34];
                mwm.DetailingRevenuePace = row.ColumnValue(7).ToString();


            }
            catch (WebException ex)
            {

            }

            return mwm;

        }

        public void UpdateYouMoveMeReviewsForecast(int NumberOfReviews)
        {
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "15-wXSG23VpPO95hO2KQtUpI7RlMaqflYpp1Dx_v9ooM";
                string range = "Morning Kickoff Numbers!B7";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                reader.UpdateSheet(spreadSheetId, range, NumberOfReviews.ToString());


            }
            catch (WebException ex)
            {

            }


        }

        public List<UniformData> RetrieveUniformReorderData(string sheetID, string bizUnit)
        {
            UniformData ud = new UniformData();
            List<UniformData> localList = new List<UniformData>();

            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

            string range = "";
            switch (bizUnit)
            {
                case "gj":
                    range = "junk!A4:D100";
                    break;

                case "ym":
                    range = "move!A4:D100";
                    break;

                case "ss":
                    range = "shack!A4:D100";
                    break;
            }

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            Spreadsheet spreadSheet = reader.GetSpreadSheet(sheetID, range);

            //**************************************
            //*********** READ THROUGH SPREADSHEET
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                try
                {
                    //check B against C
                    if (!string.IsNullOrEmpty(row.ColumnValue(1).ToString().Trim().ToLower()) && !string.IsNullOrEmpty(row.ColumnValue(2).ToString().Trim().ToLower()) && !string.IsNullOrEmpty(row.ColumnValue(3).ToString().Trim().ToLower()) && row.ColumnValue(1).ToString().Trim().ToLower() != "n/a")
                    {
                        if (float.Parse(row.ColumnValue(1).ToString().Trim()) <= float.Parse(row.ColumnValue(2).ToString().Trim()) && float.Parse(row.ColumnValue(3).ToString().Trim()) > 0)
                        {
                            //stock is less than reorder and reorder amount is greater than 0
                            ud = new UniformData();
                            ud.ItemName = "  " + row.ColumnValue(0).ToString().Trim();
                            ud.ItemStock = row.ColumnValue(1).ToString().Trim();
                            ud.ItemReorder = row.ColumnValue(2).ToString().Trim();
                            ud.ItemReorderAmount = row.ColumnValue(3).ToString().Trim();
                            localList.Add(ud);

                        }
                    }

                }
                catch (WebException ex)
                {

                }



            }

            return localList;
        }

        public string ConnectSheetsMoveBagDropROI()
        {
            string output = "*Bag Drop ROI Rolling 90 Day*\n";

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1dJW-XRk_n6d5X_OhXWHp26CZQutV2HUC4knE1aPEVyE";
                string range = "Reporting!A1:E65";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Start Date
                SpreadSheetRow row = spreadSheet.Rows[14];
                output = output + "  Start Date:  " + row.ColumnValue(2).ToString() + "\n";

                //End Date
                row = spreadSheet.Rows[15];
                output = output + "  End Date:  " + row.ColumnValue(2).ToString() + "\n\n";

                //Total Bags Dropped
                row = spreadSheet.Rows[18];
                output = output + "  Bags Dropped:  " + row.ColumnValue(2).ToString() + "\n";

                //Addresses Match
                row = spreadSheet.Rows[19];
                output = output + "  Addresses Matched:  " + row.ColumnValue(2).ToString() + "\n";

                //Address Match %
                row = spreadSheet.Rows[20];
                output = output + "  Matched %:  " + row.ColumnValue(2).ToString() + "\n\n";

                //RoadWarrior Cost
                row = spreadSheet.Rows[22];
                output = output + "  Roadwarrior Cost:  " + row.ColumnValue(2).ToString() + "\n";

                //Jobs Total @ Bags Dropped
                row = spreadSheet.Rows[23];
                output = output + "  Jobs Total @ Addresses:  " + row.ColumnValue(2).ToString() + "\n";

                //AJS Jobs Total @ Bags Dropped
                row = spreadSheet.Rows[24];
                output = output + "  AJS @ Addresses:  " + row.ColumnValue(2).ToString() + "\n";

                //ROI RoadWarrior
                row = spreadSheet.Rows[25];
                output = output + "  Roadwarrior ROI:  " + row.ColumnValue(2).ToString() + "\n";



                output = output + "\n*Bags Dropped Last 7 Days*\n";

                //Total Bags Dropped
                row = spreadSheet.Rows[33];
                output = output + "  Total Bags Dropped:  " + row.ColumnValue(2).ToString() + "\n";

                //Total Hours Worked
                row = spreadSheet.Rows[31];
                output = output + "  Total Hours Worked:  " + row.ColumnValue(2).ToString() + "\n";

                //Bags Dropped / Hour
                row = spreadSheet.Rows[34];
                output = output + "  Bags Dropped / Hour:  " + row.ColumnValue(2).ToString() + "\n\n";



                output = output + "\n*Employee - Bags Dropped - Hours Worked Last 7 Days - Dropped / Hour*\n";

                //List of Employees
                for (int x = 37; x < 61; x++) 
                {
                    row = spreadSheet.Rows[x];
                    if (!(row.ColumnValue(1).ToString().Trim() == ""))
                    {
                        output = output + "  " + row.ColumnValue(1).ToString() + " - " + row.ColumnValue(2).ToString() + " - " + row.ColumnValue(3).ToString() + " - " + row.ColumnValue(4).ToString() + "\n";
                    }
                }

            }
            catch (WebException ex)
            {

            }

            return output;

        }


        public String ConnectSheetsMoveMorningMeetingTime()
        {
            string meetingTime = "";

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs";
                string range = "Daily Schedule!F1:F2";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //get Start Date
                SpreadSheetRow row = spreadSheet.Rows[0];
                meetingTime = row.ColumnValue(0).ToString();

            }
            catch (WebException ex)
            {

            }

            return meetingTime;

        }


        public MoveDailyWages ConnectSheetsShackDailyOpsChecklist(string sheetID)
        {
            int counter = 0;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string range = "Daily Checklist!A1:H40";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(sheetID, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET


                //Manager
                m.RPH = spreadSheet.Rows[0].ColumnValue(1).ToString().Trim();               //Manager Name
                m.WagePercentage = spreadSheet.Rows[0].ColumnValue(6).ToString().Trim();    //Date
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();



                //VANS HEADLINE
                m.RouteDate = "*VANS*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                //VANS
                for (int vanCount = 3; vanCount < 8; vanCount++)
                {
                    //if (spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim().ToLower() == "not done" || spreadSheet.Rows[vanCount].ColumnValue(2).ToString().Trim().ToLower() == "false")
                    //{
                        //nobody entered then log
                        m.RouteDate = "Van " + spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim() + ": " + spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim() + "   -   Cleaned: " + spreadSheet.Rows[vanCount].ColumnValue(2).ToString().Trim() + "   -   Gas: " + spreadSheet.Rows[vanCount].ColumnValue(3).ToString().Trim() + " -    Notes: " + spreadSheet.Rows[vanCount].ColumnValue(4).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    //}
                }



                //Pre Meeting HEADLINE
                m.RouteDate = "\n*Pre Meeting*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                for (int vanCount = 9; vanCount < 13; vanCount++)
                {
                    if (spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }

                    if (spreadSheet.Rows[vanCount].ColumnValue(4).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(5).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }
                }




                //Meeting HEADLINE
                m.RouteDate = "\n*Meeting*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                for (int vanCount = 14; vanCount < 16; vanCount++)
                {
                    if (spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }

                    if (spreadSheet.Rows[vanCount].ColumnValue(4).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(5).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }
                }




                //Daily - AM HEADLINE
                m.RouteDate = "\n*Daily - AM*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                for (int vanCount = 17; vanCount < 22; vanCount++)
                {
                    if (spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }

                    if (spreadSheet.Rows[vanCount].ColumnValue(4).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(5).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }
                }



                //Daily - PM HEADLINE
                m.RouteDate = "\n*Daily - PM*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                for (int vanCount = 23; vanCount < 25; vanCount++)
                {
                    if (spreadSheet.Rows[vanCount].ColumnValue(0).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }

                    if (spreadSheet.Rows[vanCount].ColumnValue(4).ToString().Trim().ToLower() == "not done")
                    {
                        //nobody entered then log
                        m.RouteDate = "ND - " + spreadSheet.Rows[vanCount].ColumnValue(5).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();

                    }
                }


                //NOTES HEADLINE
                m.RouteDate = "\n*NOTES*\n";
                mdw.DailyWages.Add(m);
                m = new MoveDailyWage();

                for (int vanCount = 26; vanCount < 32; vanCount++)
                {
                    if (!(spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim() == ""))
                    {
                        m.RouteDate = spreadSheet.Rows[vanCount].ColumnValue(1).ToString().Trim();
                        mdw.DailyWages.Add(m);
                        m = new MoveDailyWage();
                    }
                }




            }
            catch (WebException ex)
            {

            }

            return mdw;


        }

        public MoveDailyWage ConnectSheetsGetOpenAR(string sheetID, string range)
        {
            MoveDailyWage m = new MoveDailyWage();
            m.RPH = "";

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);
                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(sheetID, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        m.RPH = row.ColumnValue(1).ToString().Trim().Replace("-", "");
                    }
                    catch (WebException ex)
                    {

                    }
                }

            }
            catch (WebException ex)
            {

            }

            if (m.RPH.Trim() == "")
            {
                m.RPH = "ERROR: open amount not found in cell J1";
            }

            return m;
        }



        public MoveDailyWages ConnectSheetsDailyChecklist(string sheetID)
        {
            int counter = 2;
            MoveDailyWage m = new MoveDailyWage();
            MoveDailyWages mdw = new MoveDailyWages();
            mdw.DailyWages = new List<MoveDailyWage>();


            //Get the Data in the current sheet
            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                //string spreadSheetId = "1rMFYFIfiuKe6si-10CTHXR1tUMbNMYnD7nJ__9V-PBU";          //YMM Daily Operations Checklist
                string range = "Daily Checklist!A1:D200";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(sheetID, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if(row.ColumnValue(3).ToString().Trim() == "1")
                        {
                            m.RouteDate = row.ColumnValue(0).ToString().Trim();
                            m.RouteNumber = row.ColumnValue(1).ToString().Trim().Replace("'","");
                            m.WagePercentage = counter.ToString();

                            if(m.RouteDate.ToString().Trim() == "")
                            {
                                m.RouteDate = "*MISSING*";
                            }

                            if (m.RouteDate.ToString().Trim().ToLower() == "not done")
                            {
                                m.RouteDate = "*NOT DONE*";
                            }

                            mdw.DailyWages.Add(m);
                            m = new MoveDailyWage();
                        }
                        else if (row.ColumnValue(3).ToString().Trim() == "99")
                        {
                            m.RPH = row.ColumnValue(0).ToString().Trim();
                            m.WagePercentage = "x";
                            mdw.DailyWages.Add(m);
                            m = new MoveDailyWage();
                        }

                        counter = counter + 1;


                    }
                    catch (WebException ex)
                    {

                    }
                }

            }
            catch (WebException ex)
            {

            }

            return mdw;


        }

        public MoveDailyWage ConnectSheetsMoveSalesCenterWages(DateTime startDate, DateTime endDate)
        {
            DateTime temp;
            int dateDiff = (endDate - startDate).Days + 1;

            string scLeadWages = "0";
            string scHourlyWages = "0";
            float hourlyWageTaylor = 0;
            float hourlyWageJake = 0;
            MoveDailyWage m = new MoveDailyWage();

            string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
            string applicationName = "My Project";
            string[] scopes = { SheetsService.Scope.Spreadsheets };

            GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);



            //get scLeadWages - Jacob via the Sales Center Scheduling sheet
            string spreadSheetId = "1MTwqtKym02eP6xHX3OmmbnfNtxG9CocS_YCTrQ_rqiM";          //Sales Center Scheduling
            string range = "SC Staffing Budget!A1:D30";

            GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
            Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            scLeadWages = (float.Parse(spreadSheet.Rows[9].ColumnValue(2).Replace("$", "").Replace(",", "")) / 30 * dateDiff).ToString();


            //get hourly wage for Taylor and Jake
            spreadSheetId = "1MTwqtKym02eP6xHX3OmmbnfNtxG9CocS_YCTrQ_rqiM";          //Sales Center Scheduling
            range = "SC Slack Reporting!A1:D30";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            hourlyWageTaylor = float.Parse(spreadSheet.Rows[11].ColumnValue(3).Replace("$", "").Replace(",", ""));
            hourlyWageJake = float.Parse(spreadSheet.Rows[11].ColumnValue(2).Replace("$", "").Replace(",", ""));



            //get scHourlyWages - Haylee
            //spreadSheetId = "1DiUMtGz37HVqk9jGEQNIVJZAiMQa33BtJchzaf-KgxY";          //Taylor TimeCard
            spreadSheetId = "12q0U8KxM1Y2bVMeJT2usxVTILLmwcEnXHlXQGsXemvI";         //Haylee Timecard
            range = "HOURS SUMMARY!A1:F500";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //**************************************
            //*********** READ THROUGH SPREADSHEET
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                try
                {
                    DateTime.TryParse(row.ColumnValue(0).ToString().Trim(), out temp);
                    if (temp >= startDate && temp <= endDate)
                    {
                        if (row.ColumnValue(4).ToString().ToLower().Trim() == "ymm: sales center")
                        {
                            scHourlyWages = (float.Parse(scHourlyWages) + (float.Parse(row.ColumnValue(3).ToString()) * hourlyWageTaylor)).ToString();
                        }
                    }

                }
                catch (WebException ex)
                {

                }
            }




            //get scHourlyWages - Jake
            spreadSheetId = "1PLl4LBAZ2ov0ebt4RSnv0Vh39_QGQV01fPZTb1Vr5ss";          //Jake TimeCard
            range = "HOURS SUMMARY!A1:M500";

            reader = new GoogleSpreadSheetReader(googleService);
            spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

            //**************************************
            //*********** READ THROUGH SPREADSHEET
            foreach (SpreadSheetRow row in spreadSheet.Rows)
            {
                try
                {
                    DateTime.TryParse(row.ColumnValue(0).ToString().Trim(), out temp);
                    if (temp >= startDate && temp <= endDate)
                    {
                        if (row.ColumnValue(4).ToString().ToLower().Trim() == "ymm: sales center" && !(row.ColumnValue(3).ToString().Trim() == "-"))
                        {
                            scHourlyWages = (float.Parse(scHourlyWages) + (float.Parse(row.ColumnValue(3).ToString()) * hourlyWageJake)).ToString();
                        }
                    }

                    //DateTime.TryParse(row.ColumnValue(0).ToString().Trim(), out temp);
                    //if (temp >= startDate && temp <= endDate)
                    //{
                    //    scHourlyWages = (float.Parse(scHourlyWages) + (float.Parse(row.ColumnValue(11).ToString()) * hourlyWageJake)).ToString();
                    //}

                }
                catch (WebException ex)
                {

                }
            }








            m.RouteDate = (float.Parse(scLeadWages) + float.Parse(scHourlyWages)).ToString();

            return m;
        }


        public MoveDailyWage ConnectSheetsSalesWages()
        {
            MoveDailyWage m = new MoveDailyWage();

            try
            {

                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1MTwqtKym02eP6xHX3OmmbnfNtxG9CocS_YCTrQ_rqiM";          //Sales Center Scheduling
                string range = "sales_center_reporting!A1:C11";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //YMM Wage%
                m.MiscItem = spreadSheet.Rows[5].ColumnValue(2).ToString().Trim();

                //YMM Goal%
                m.MiscItem2 = spreadSheet.Rows[6].ColumnValue(2).ToString().Trim();

                //YMM O/U %
                m.MiscItem3 = spreadSheet.Rows[7].ColumnValue(2).ToString().Trim();

                //YMM O/U $
                m.MiscItem4 = spreadSheet.Rows[8].ColumnValue(2).ToString().Trim();

                //SS Wage%
                m.MiscItem5 = spreadSheet.Rows[5].ColumnValue(1).ToString().Trim();

                //SS Goal%
                m.MiscItems6 = spreadSheet.Rows[6].ColumnValue(1).ToString().Trim();

                //SS O/U %
                m.CrewLead = spreadSheet.Rows[7].ColumnValue(1).ToString().Trim();

                //SS O/U $
                m.EndDate = spreadSheet.Rows[8].ColumnValue(1).ToString().Trim();

            }
            catch (WebException ex)
            {

            }

            return m;

        }


    public MoveDailyWage ConnectSheetsMoveDailyScorecard(DateTime startDate, DateTime endDate)
        {
            int dataIndex = 0;
            DateTime temp;
            bool startCounting = false;

            string totalRevenue = "0";
            string totalWages = "0";
            string totalMovingHours = "0";
            int totalAJS = 0;

            MoveDailyWage m = new MoveDailyWage();
            DateTime now = DateTime.Now;

            //the intent is to cycle through the sheet and get the total revenue for the day before 


            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                //string spreadSheetId = "10kWJZPFUIlgMxsSzCxdtTPy6f79a6bKJuILR_uwE7UI";          //YMM Job Tracking 
                string spreadSheetId = "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I";          //YMM Payroll & Labor
                string range = "Job Tracking!A1:BK8000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        DateTime.TryParse(row.ColumnValue(0).ToString().Trim(), out temp);
                        if (temp >= startDate && temp <= endDate) 
                        {
                            startCounting = true;
                        }
                        if (startCounting)
                        {
                            if (dataIndex == 0)
                            {
                                //Moving hours
                                if (row.ColumnValue(58).ToString().IndexOf("-") == -1)
                                {
                                    totalMovingHours = (float.Parse(totalMovingHours) + float.Parse(row.ColumnValue(58).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }

                                dataIndex = dataIndex + 1;
                            }
                            else if (dataIndex == 5)
                            {
                                //Total Sales
                                if (row.ColumnValue(56).ToString().IndexOf("-") == -1)
                                {
                                    totalRevenue = (float.Parse(totalRevenue) + float.Parse(row.ColumnValue(56).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }

                                //Total Wages     
                                if (row.ColumnValue(59).ToString().IndexOf("-") == -1)
                                {
                                    totalWages = (float.Parse(totalWages) + float.Parse(row.ColumnValue(59).ToString().Trim().Replace("$", "").Replace(" ", ""))).ToString();
                                }


                                startCounting = false;
                                dataIndex = 0;
                                totalAJS = totalAJS + 1;
                            }
                            else
                            {
                                dataIndex = dataIndex + 1;
                            }
                        }

                    }
                    catch (WebException ex)
                    {

                    }

                }


                //get data from the Wage Management tab
                range = "Wage Management!A1:B100";

                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Wage%
                m.MiscItem = spreadSheet.Rows[5].ColumnValue(1).ToString().Trim();

                //Wage% goal
                m.MiscItem2 = spreadSheet.Rows[6].ColumnValue(1).ToString().Trim();

                //OH%
                m.MiscItem3 = spreadSheet.Rows[91].ColumnValue(1).ToString().Trim();

                //OH% goal
                m.MiscItem4 = spreadSheet.Rows[92].ColumnValue(1).ToString().Trim();

            }
            catch (WebException ex)
            {

            }

            //Revenue
            m.RouteDate = totalRevenue.ToString();

            //Direct Wage %
            m.RouteNumber = (float.Parse(totalWages.Replace("$", "").Replace(" ","")) / float.Parse(totalRevenue.Replace("$", "").Replace(" ", ""))).ToString();

            //AJS
            m.WagePercentage = (float.Parse(totalRevenue.Replace("$", "").Replace(" ", "")) / totalAJS).ToString();

            //RPH
            m.RPH = (float.Parse(totalRevenue.Replace("$", "").Replace(" ", "")) / float.Parse(totalMovingHours.Replace("$", "").Replace(" ", ""))).ToString();









            return m;

        }












        public MoveDailyWages ConnectSheetsMoveDailyWage()
        {
            int dataIndex = 0;
            bool startCounting = false;

            MoveDailyWages mdw = new MoveDailyWages();
            MoveDailyWage m = new MoveDailyWage();
            mdw.DailyWages = new List<MoveDailyWage>();

            DateTime now = DateTime.Now;

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "10kWJZPFUIlgMxsSzCxdtTPy6f79a6bKJuILR_uwE7UI";          //YMM Job Tracking 
                //string spreadSheetId = "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I";          //YMM Payroll & Labor
                string range = "Job Tracking!A1:BK1000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() == now.AddDays(-1).ToShortDateString())
                        {
                            startCounting = true;

                            //BK = 62
                            if (!string.IsNullOrEmpty(row.ColumnValue(62).ToString().Trim()))
                            {
                                m.RouteDate = row.ColumnValue(1).ToString().Trim(); ;
                                m.RouteNumber = row.ColumnValue(2).ToString().Trim(); ;
                                m.CrewLead = row.ColumnValue(62).ToString().Trim();
                            }
                        }

                        if (startCounting)
                        {
                            if (dataIndex == 6)
                            {
                                //found the row

                                //BI = 60
                                m.WagePercentage = row.ColumnValue(60).ToString().Trim();

                                //BF = 57
                                m.RPH = row.ColumnValue(57).ToString().Trim();

                                startCounting = false;
                                dataIndex = 0;

                                mdw.DailyWages.Add(m);
                                m = new MoveDailyWage();

                            }
                            else
                            {
                                dataIndex = dataIndex + 1;
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

            return mdw;

        }








        public ShackDailyWages ConnectSheetsShackDailyWage()
        {
            int dataIndex = 0;
            bool startCounting = false;

            ShackDailyWages sdw = new ShackDailyWages();
            ShackDailyWage m = new ShackDailyWage();
            sdw.DailyWages = new List<ShackDailyWage>();

            DateTime now = DateTime.Now;

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1ttCMCGo-C5ozvlgWgMvQKkSY0Ujw4hO_yJlL3qLOn_Q";
                string range = "Job Tracking!A1:BT5000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //**************************************
                //*********** READ THROUGH SPREADSHEET
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() == now.AddDays(-1).ToShortDateString())
                        {
                            startCounting = true;

                            m.RouteDate = row.ColumnValue(1).ToString().Trim(); ;
                            m.RouteNumber = row.ColumnValue(2).ToString().Trim(); ;
                            m.CrewLead = "";
                        }

                        if (startCounting)
                        {
                            if(dataIndex == 4)
                            {
                                //BQ = 68
                                m.RPH = row.ColumnValue(68).ToString().Trim();

                                dataIndex = dataIndex + 1;
                            }
                            else if (dataIndex == 6)
                            { 
                                //BT = 71
                                m.WagePercentage = row.ColumnValue(71).ToString().Trim();
                                if (m.WagePercentage.Contains("DIV"))
                                {
                                    m.WagePercentage = "0%";
                                }

                                startCounting = false;
                                dataIndex = 0;

                                sdw.DailyWages.Add(m);
                                m = new ShackDailyWage();

                            }
                            else if (dataIndex == 0)
                            {
                                //BQ = 68
                                m.Revenue = row.ColumnValue(68).ToString().Trim();

                                dataIndex = dataIndex + 1;
                            }
                            else
                            {
                                dataIndex = dataIndex + 1;
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

            return sdw;

        }


        public List<Employee> ConnectSheetsBuildingSecurity()
        {

            List<Employee> empList = new List<Employee>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1Lcd5o8sOq-lkxoPATiZHykwbPU7aNuyJLXVVUlbWPqc";
                string range = "Daily Leadership Plan!H1:M2";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheetNoHeader(spreadSheetId, range);

                //Column H - Date
                Employee e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(0).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(0).ToString().Trim();

                empList.Add(e);

                //Column I - Morning Greeter
                e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(1).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(1).ToString().Trim();

                empList.Add(e);

                //Column J - Morning Greeter
                e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(2).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(2).ToString().Trim();

                empList.Add(e);

                //Column K - Morning Uniforms
                e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(3).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(3).ToString().Trim();

                empList.Add(e);

                //Column L - AM Magic
                e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(4).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(4).ToString().Trim();

                empList.Add(e);

                //Column M - Close
                e = new Employee();
                e.EmpName = spreadSheet.Rows[0].ColumnValue(5).ToString().Trim();
                e.EmpPhone = spreadSheet.Rows[1].ColumnValue(5).ToString().Trim();

                empList.Add(e);

                ////Column T - GJ AM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(6).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(6).ToString().Trim();

                //empList.Add(e);

                ////Column U - GJ PM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(7).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(7).ToString().Trim();

                //empList.Add(e);

                ////Column V - SS AM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(8).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(8).ToString().Trim();

                //empList.Add(e);

                ////Column W - SS PM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(9).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(9).ToString().Trim();

                //empList.Add(e);

                ////Column X - YM AM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(10).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(10).ToString().Trim();

                //empList.Add(e);

                ////Column Y - YM PM Point
                //e = new Employee();
                //e.EmpName = spreadSheet.Rows[0].ColumnValue(11).ToString().Trim();
                //e.EmpPhone = spreadSheet.Rows[1].ColumnValue(11).ToString().Trim();

                //empList.Add(e);

            }
            catch (WebException ex)
            {

            }

            return empList;


        }


        public List<JunkDailySchedule> ConnectSheetsJunkDailySchedule(string sheetName, string rangeName)
        {
            string range = "";
            List<Employee> empList = new List<Employee>();
            List<JunkDailySchedule> localList = new List<JunkDailySchedule>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ";
                range = sheetName + rangeName;

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and populate a list of employees
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(0).ToString().Trim() != "")
                        {
                            Employee e = new Employee();
                            e.EmpName = row.ColumnValue(0).ToString().Trim();
                            e.EmpPhone = row.ColumnValue(1).ToString().Trim().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");

                            empList.Add(e);
                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }



                range = sheetName + "A1:D8";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Point
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(0).ToString().Trim() != "")
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(0).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = "Point";

                                    localList.Add(jds);

                                }
                            }

                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }




                range = sheetName + "A9:D14";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Point
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() != "")
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(1).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = "On-Call";

                                    localList.Add(jds);

                                }
                            }

                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }




                range = sheetName + "A14:D100";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Point
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() != "" & !row.ColumnValue(1).ToString().Contains("Truck"))
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(1).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = row.ColumnValue(2).ToString();

                                    localList.Add(jds);

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

            return localList;

        }



        public List<JunkDailySchedule> ConnectSheetsSCAfterHoursLeads()
        {

            TimeSpan startAfternoon = new TimeSpan(0, 0, 0); //Midnight
            TimeSpan endAfternoon = new TimeSpan(23, 59, 59); //11:59:59 pm
            TimeSpan startMorning = new TimeSpan(0, 0, 0); //Midnight
            TimeSpan endMorning = new TimeSpan(23, 59, 59); //11:59:59 pm
            TimeSpan now = DateTime.Now.TimeOfDay;

            //only run on this schedule
            //Sunday 3pm - 10am
            //Monday - Friday 6pm - 8am
            //Saturday 5pm - 9am

            DateTime timeTemp = DateTime.Now;











            List<Employee> empList = new List<Employee>();
            List<JunkDailySchedule> localList = new List<JunkDailySchedule>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1T_htIBFPsNe2xCgXFvgUdSec3s_-nupKs4S204HfTr4";
                string range = "OBE Leads!A1:D5000";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Team Lead
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        DateTime outTemp;
                        DateTime.TryParse(row.ColumnValue(0).ToString().Trim(), out outTemp);

                        if (DateTime.Now.ToShortDateString() == outTemp.ToShortDateString())
                        {
                            now = outTemp.TimeOfDay;
                            switch (timeTemp.DayOfWeek.ToString())
                            {
                                case "Sunday":
                                    endMorning = new TimeSpan(9, 59, 59); //9:59:59 am
                                    startAfternoon = new TimeSpan(15, 0, 0); //3:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Monday":
                                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Tuesday":
                                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Wednesday":
                                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Thursday":
                                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Friday":
                                    endMorning = new TimeSpan(7, 59, 59); //7:59:59 am
                                    startAfternoon = new TimeSpan(18, 0, 0); //6:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
                                case "Saturday":
                                    endMorning = new TimeSpan(8, 59, 59); //8:59:59 am
                                    startAfternoon = new TimeSpan(17, 0, 0); //5:00:00 pm

                                    if ((((now > startMorning) && (now < endMorning)) || (((now > startAfternoon) && (now < endAfternoon)))))
                                    {
                                        JunkDailySchedule jds = new JunkDailySchedule();
                                        jds.EmpName = row.ColumnValue(1).ToString().Trim();
                                        jds.EmpPhone = row.ColumnValue(3).ToString().Trim();

                                        localList.Add(jds);
                                    }
                                    break;
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

            return localList;

        }



        public List<JunkDailySchedule> ConnectSheetsMoveDailySchedule()
        {
            List<Employee> empList = new List<Employee>();
            List<JunkDailySchedule> localList = new List<JunkDailySchedule>();

            try
            {
                string googleSecretJsonFilePath = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\Content\\GoogleSecret.json";
                string applicationName = "My Project";
                string[] scopes = { SheetsService.Scope.Spreadsheets };

                GoogleService googleService = new GoogleService(googleSecretJsonFilePath, applicationName, scopes);

                string spreadSheetId = "1LQn5oRmairJUEgh7_qEOi-gSp8wrN8eVCVqXfwBG3bs";
                string range = "YMM Team!AB2:AE100";

                GoogleSpreadSheetReader reader = new GoogleSpreadSheetReader(googleService);
                Spreadsheet spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //JunkDailySchedule jds = new JunkDailySchedule();
                //jds.EmpName = "Andrew";
                //jds.EmpPhone = "6122813895";
                //jds.EmpStartTime = "9:00 AM";
                //jds.EmpWorkType = "Point";

                //localList.Add(jds);


                //Read through and populate a list of employees
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(0).ToString().Trim() != "")
                        {
                            Employee e = new Employee();
                            e.EmpName = row.ColumnValue(0).ToString().Trim();
                            e.EmpPhone = row.ColumnValue(3).ToString().Trim().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");

                            empList.Add(e);
                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }



                range = "Daily Schedule!A5:D8";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Team Lead
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(0).ToString().Trim() != "")
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(0).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = "Team Lead";

                                    localList.Add(jds);

                                }
                            }

                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }




                range = "Daily Schedule!A9:D13";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get On-call
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() != "")
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(1).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = "On-Call";

                                    localList.Add(jds);

                                }
                            }

                        }
                    }
                    catch (WebException ex)
                    {

                    }
                }



                string route = "";
                range = "Daily Schedule!A13:E100";
                reader = new GoogleSpreadSheetReader(googleService);
                spreadSheet = reader.GetSpreadSheet(spreadSheetId, range);

                //Read through and get Team Schedule
                foreach (SpreadSheetRow row in spreadSheet.Rows)
                {
                    try
                    {
                        if (row.ColumnValue(1).ToString().Trim() != "" & !row.ColumnValue(4).ToString().Trim().ToLower().Contains("truck"))
                        {

                            foreach (Employee emp in empList)
                            {
                                if (emp.EmpName == row.ColumnValue(1).ToString().Trim())
                                {
                                    JunkDailySchedule jds = new JunkDailySchedule();
                                    jds.EmpName = emp.EmpName;
                                    jds.EmpPhone = emp.EmpPhone;
                                    jds.EmpStartTime = row.ColumnValue(3).ToString();
                                    jds.EmpWorkType = route;

                                    localList.Add(jds);

                                }
                            }

                        }
                        else if (row.ColumnValue(0).ToString().Trim().Contains("RT"))
                        {
                            route = row.ColumnValue(0).ToString().Trim();
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

            return localList;

        }











    }
}