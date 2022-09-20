using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text;
using Utilities360Wow;
using NightlyRouteToSlack.Utilities;

using Data = Google.Apis.Sheets.v4.Data;

namespace NightlyRouteToSlack 
{
    class GoogleSpreadSheetReader
    {
        private readonly SheetsService _sheetService;
        public GoogleSpreadSheetReader(GoogleService googleService)
        {
            _sheetService = googleService.GetSheetsService();
        }

        public Spreadsheet GetSpreadSheet(string spreadSheetId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(spreadSheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                var row = new SpreadSheetRow(values[i]);
                rows.Add(row);
            }
            var headerRow = new SpreadSheetRow(values[0]);
            var spreadSheet = new Spreadsheet();
            spreadSheet.HeaderRow = headerRow;
            spreadSheet.Rows = new List<SpreadSheetRow>();
            spreadSheet.Rows.AddRange(rows);
            return spreadSheet;
        }

        public Spreadsheet GetSpreadSheetNoHeader(string spreadSheetId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(spreadSheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 0; i < values.Count; i++)
            {
                var row = new SpreadSheetRow(values[i]);
                rows.Add(row);
            }
            var headerRow = new SpreadSheetRow(values[0]);
            var spreadSheet = new Spreadsheet();
            spreadSheet.HeaderRow = headerRow;
            spreadSheet.Rows = new List<SpreadSheetRow>();
            spreadSheet.Rows.AddRange(rows);
            return spreadSheet;
        }

        public void UpdateSheet(string sheetID, string rangeToUpdate, string numberOfReviews)
        {
            string spreadsheetId = sheetID;  

            string range = rangeToUpdate;  

            SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = (SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum)2;  

            Data.ValueRange requestBody = new Data.ValueRange();

            var test = new string[] { numberOfReviews };
            requestBody.Values = new List<IList<object>> { test };

            SpreadsheetsResource.ValuesResource.UpdateRequest request = _sheetService.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
            request.ValueInputOption = valueInputOption;

            Data.UpdateValuesResponse response = request.Execute();

            

        }

        private int ReturnLastRowOfData(string sheetID)
        {
            int valueToReturn = 3;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkRouteJournalLastRow"));

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;
        }

        private int ReturnLastRowOfDataMoveTips(string sheetID, string tabName, int valueToReturn)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, tabName);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;

        }

        private int ReturnLastRowOfDataMove(string sheetID)
        {
            int valueToReturn = 10;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveRouteJournalLastRow"));

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;
        }

        private int ReturnLastRowOfDataEmployeeHoursWorked(string sheetID)
        {
            int valueToReturn = 2;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, "Employee Hours!A1:A");

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;

        }


        private int ReturnLastRowOfDataEmployeeJobDone(string sheetID)
        {
            int valueToReturn = 2;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, "GJ Data Dump!A1:A");

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;

        }

        private int ReturnLastRowOfDataHealthCheck(string sheetID)
        {
            int valueToReturn = 2;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, "DailyCheckAll!A1:A");

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;

        }

        private int ReturnLastRowOfDataShack(string sheetID)
        {
            int valueToReturn = 3;
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetService.Spreadsheets.Values.Get(sheetID, cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackRouteJournalLastRow"));

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            var rows = new List<SpreadSheetRow>();
            for (int i = 1; i < values.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(values[i][0].ToString()))
                    {
                        valueToReturn = valueToReturn + 1;
                    }
                }
                catch
                {
                    break;
                }
            }

            return valueToReturn;
        }






        public void UpdateShackDailyOpsChecklist(string sheetID)
        {
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj = new List<Object>();
            obj.Add("Not Done");
            objNewRecords.Add(obj);

            List<IList<Object>> objNewRecordsFALSE = new List<IList<Object>>();
            IList<Object> objFALSE = new List<Object>();
            objFALSE.Add("FALSE");
            objNewRecordsFALSE.Add(objFALSE);

            List<IList<Object>> objNewRecordsEMPTY = new List<IList<Object>>();
            IList<Object> objEMPTY = new List<Object>();
            objEMPTY.Add("");
            objNewRecordsEMPTY.Add(objEMPTY);

            //Manager Name
            string rangeToUpdate = "Daily Checklist!B2:B2";

            SpreadsheetsResource.ValuesResource.UpdateRequest request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();


            //First Section (VANS)
            for (int y = 5; y < 11; y++)
            {
                rangeToUpdate = "Daily Checklist!A" + y.ToString() + ":A" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();


                rangeToUpdate = "Daily Checklist!C" + y.ToString() + ":C" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecordsFALSE }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

                rangeToUpdate = "Daily Checklist!E" + y.ToString() + ":E" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecordsEMPTY }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();
            }


            //Second Section (Pre Meeting)
            for (int y = 12; y < 16; y++)
            {
                rangeToUpdate = "Daily Checklist!A" + y.ToString() + ":A" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

                rangeToUpdate = "Daily Checklist!E" + y.ToString() + ":E" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();
            }

            //Third Section (Meeting)
            for (int y = 17; y < 19; y++)
            {
                rangeToUpdate = "Daily Checklist!A" + y.ToString() + ":A" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

                rangeToUpdate = "Daily Checklist!E" + y.ToString() + ":E" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();
            }


            //Fourth Section (Daily - AM)
            for (int y = 20; y < 25; y++)
            {
                rangeToUpdate = "Daily Checklist!A" + y.ToString() + ":A" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

                rangeToUpdate = "Daily Checklist!E" + y.ToString() + ":E" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();
            }

            //Fifth Section (Daily - PM)
            for (int y = 26; y < 28; y++)
            {
                rangeToUpdate = "Daily Checklist!A" + y.ToString() + ":A" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

                rangeToUpdate = "Daily Checklist!E" + y.ToString() + ":E" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();
            }


            //Notes Section
            for (int y = 29; y < 35; y++)
            {
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add("");
                objNewRecords.Add(obj);
                rangeToUpdate = "Daily Checklist!B" + y.ToString() + ":B" + y.ToString();

                request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();

            }


        }


        public void UpdateMoveDailyOpsChecklist(string sheetID, MoveDailyWages m)
        {
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj = new List<Object>();
            obj.Add("Not Done");
            objNewRecords.Add(obj);

            foreach (MoveDailyWage mdw in m.DailyWages)
            {
                if (!(mdw.WagePercentage == "x"))
                {
                    string rangeToUpdate = "Daily Checklist!A" + mdw.WagePercentage + ":A" + mdw.WagePercentage;

                    SpreadsheetsResource.ValuesResource.UpdateRequest request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    var response = request.Execute();

                }

            }

        }

        public void UpdateJunkHoursWorked(string sheetID, DataTable dt)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            string worksheetName = "raw_data";
            int initialRow = ReturnLastRowOfDataMoveTips(sheetID, "raw_data!A1:A", 1);
            int counter = initialRow;

            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            foreach (DataRow dr in dt.Rows)
            {
                IList<Object> obj = new List<Object>();
                obj.Add(dr["routeID"].ToString());
                obj.Add(dr["userID"].ToString());
                obj.Add(dr["firstName"].ToString() + " " + dr["lastName"].ToString());
                obj.Add(dr["startTime"].ToString());
                obj.Add(dr["endTime"].ToString());
                obj.Add(dr["hours"].ToString());
                obj.Add(dr["wage"].ToString());
                obj.Add((float.Parse(dr["hours"].ToString()) * float.Parse(dr["wage"].ToString())).ToString());
                obj.Add(dr["workType"].ToString());
                obj.Add(DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                objNewRecords.Add(obj);
                counter = counter + 1;
            }

            string rangeToUpdate = worksheetName + "!A" + initialRow + ":I" + counter;

            SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
        }


        public void UpdateJunkSquareTransactions(string sheetID, List<SquareData> localList)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            string worksheetName = cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkRouteJournalLastRow").ToString();
            worksheetName = worksheetName.Substring(0, worksheetName.IndexOf("!"));
            int counter = ReturnLastRowOfData(sheetID);

            foreach (SquareData x in localList)
            {
                List<IList<Object>> objNewRecords = new List<IList<Object>>();
                IList<Object> obj = new List<Object>();
                obj.Add(x.RouteDate);
                obj.Add(x.Route);
                obj.Add(x.JobID);
                obj.Add(x.Revenue);
                obj.Add(x.MattressCount);
                obj.Add(x.MattressValue);
                obj.Add(x.TVLargeCount);
                obj.Add(x.TVLargeValue);
                obj.Add(x.TVSmallCount);
                obj.Add(x.TVSmallValue);
                obj.Add(x.TiresCount);
                obj.Add(x.TiresValue);
                obj.Add(x.Discount);
                obj.Add(x.Tax);
                obj.Add(x.Tip);
                obj.Add(x.RMTip);
                obj.Add(x.RMMatches);
                obj.Add(x.RMLink);
                obj.Add(DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                objNewRecords.Add(obj);


                string rangeToUpdate = worksheetName + "!B" + counter + ":T" + counter;
                //string rangeToUpdate = worksheetName + "!B" + counter + ":Q" + counter;

                SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var response = request.Execute();

                counter = counter + 1;

            }


        }





        public void UploadSlackUsers(string sheetID, List<SlackUser> localList)
        {
            string rangeToUpdate = "";
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            string worksheetName = "SlackUsers";
            int counter = 2;

            foreach (SlackUser x in localList)
            {
                IList<Object> obj = new List<Object>();
                obj.Add(x.RealName);
                obj.Add(x.SlackName);
                obj.Add(x.ID);
                objNewRecords.Add(obj);

                counter = counter + 1;

            }

            rangeToUpdate = worksheetName + "!A2:C" + counter;

            SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();


        }








        public void UpdateMoveTipTransactions(string sheetID, List<SquareData> localList)
        {
            string worksheetName = "Tips";
            int counter = ReturnLastRowOfDataMoveTips(sheetID, "Tips", 1);

            foreach (SquareData x in localList)
            {
                List<IList<Object>> objNewRecords = new List<IList<Object>>();
                IList<Object> obj = new List<Object>();
                obj.Add(x.RouteDate);       //Date
                obj.Add("34.65");           //Tracker Tips
                obj.Add(x.Tip);             //Square Tips
                objNewRecords.Add(obj);

                string rangeToUpdate = worksheetName + "!A" + counter + ":C" + counter;

                SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var response = request.Execute();

                counter = counter + 1;

            }

        }






        public void UpdateMoveSquareTransactions(string sheetID, List<SquareDataMove> localList)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            string worksheetName = cs.ReturnConfigSetting("NightlyRouteToSlack", "MoveRouteJournalLastRow").ToString();
            worksheetName = worksheetName.Substring(0, worksheetName.IndexOf("!"));
            int counter = ReturnLastRowOfDataMove(sheetID);

            foreach (SquareDataMove x in localList)
            {
                List<IList<Object>> objNewRecords = new List<IList<Object>>();


                IList<Object> obj = new List<Object>();
                obj.Add(x.RouteDate);       //date
                obj.Add(x.TotalCollected);           //total collected
                obj.Add(x.PaymentType);           //payment type
                obj.Add(x.PanSuffix);         //PAN suffix
                obj.Add(x.Tip);             //tip
                obj.Add(x.Tax);             //tax
                obj.Add(x.TaxableRevenue);    //box/mat rev
                objNewRecords.Add(obj);

                string rangeToUpdate = worksheetName + "!E" + counter + ":K" + counter;

                SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var response = request.Execute();


                //CC entry type (MANUAL or SWIPED)
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add(x.SwipeOrKeyed);
                objNewRecords.Add(obj);

                rangeToUpdate = worksheetName + "!T" + counter + ":T" + counter;

                request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();


                //Imported date/time
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add(DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                objNewRecords.Add(obj);

                rangeToUpdate = worksheetName + "!S" + counter + ":S" + counter;

                request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();


                //Work Order ID
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add(x.JobID);
                objNewRecords.Add(obj);

                rangeToUpdate = worksheetName + "!P" + counter + ":P" + counter;

                request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();



                counter = counter + 1;

            }


        }


        public void UpdateJunkOTAwareness(List<Utilities.UniformData> sd)
        {


            for (int x = 4; x < 101; x++)
            {
                List<IList<Object>> objNewRecords = new List<IList<Object>>();
                IList<Object> obj = new List<Object>();

                if (x < sd.Count)
                {
                    obj.Add(sd[x].ItemName + " " + sd[x].ItemReorder);
                    obj.Add(sd[x].ItemReorderAmount);
                }
                else
                {
                    obj.Add("");
                    obj.Add("0.00");
                }

                objNewRecords.Add(obj);

                string rangeToUpdate = "OT Awareness!B" + x + ":C" + x;

                SpreadsheetsResource.ValuesResource.UpdateRequest request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, "1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ", rangeToUpdate);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var response = request.Execute();

            }

        }



        public void UpdateShackOTAwareness()
        {

            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj = new List<Object>();
            obj.Add(DateTime.Now.ToShortDateString());       //date
            objNewRecords.Add(obj);

            string rangeToUpdate = "OT Awareness!A4";


            //UPDATE Shack Shine OT Awareness Tab
            SpreadsheetsResource.ValuesResource.UpdateRequest request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, "1ttCMCGo-C5ozvlgWgMvQKkSY0Ujw4hO_yJlL3qLOn_Q", rangeToUpdate);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();

            //UPDATE You Move Me OT Awareness Tab
            request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, "1qQZUtrfX-2lPManl6VOk9Twb7vu7qa_Wi2gN6-HTH7I", rangeToUpdate);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            response = request.Execute();

            //UPDATE 1-800-GOT-JUNK? OT Awareness Tab
            request = _sheetService.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, "1xnKzcbAuJE3yYA4e9Ezjy9LESbAwlcXxCYOJAOgnTWQ", rangeToUpdate);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            response = request.Execute();

        }


        public void ImportDailyEmployeeHoursWorked(string sheetID, List<JunkDailySchedule> localList)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            int initialCounter = ReturnLastRowOfDataEmployeeHoursWorked(sheetID);
            int counter = initialCounter;

            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj = new List<Object>();

            foreach (JunkDailySchedule x in localList)
            {
                obj = new List<Object>();
                obj.Add(x.EmpStartTime);        //date
                obj.Add(x.EmpName);             //employeeName
                obj.Add(x.EmpWorkType);         //hoursWorked
                objNewRecords.Add(obj);
                counter = counter + 1;
            }

            string rangeToUpdate = "Employee Hours!A" + initialCounter + ":C" + counter;

            SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();

        }


        public void ImportDailyEmployeeJobDone(string sheetID, List<JunkDailySchedule> localList)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            int initialCounter = ReturnLastRowOfDataEmployeeJobDone(sheetID);
            int counter = initialCounter;

            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            IList<Object> obj = new List<Object>();

            foreach (JunkDailySchedule x in localList)
            {
                obj = new List<Object>();
                obj.Add(x.EmpStartTime);        //date
                obj.Add(x.EmpPhone);            //employeeID
                obj.Add(x.EmpName);             //employeeName
                obj.Add(x.EmpWorkType);         //countJobs
                objNewRecords.Add(obj);
                counter = counter + 1;
            }

            string rangeToUpdate = "GJ Data Dump!A" + initialCounter + ":D" + counter;

            SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();

        }

        public void ImportDailyHealthCheckSheet(string sheetID, List<JunkDailySchedule> localList, string biz)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            int counter = ReturnLastRowOfDataHealthCheck(sheetID);

            foreach (JunkDailySchedule x in localList)
            {

                if (!(x.EmpWorkType.ToString().Trim().ToLower() == "on-call"))
                {
                    //Tomorrow Date
                    List<IList<Object>> objNewRecords = new List<IList<Object>>();
                    IList<Object> obj = new List<Object>();
                    obj.Add(DateTime.Now.AddDays(1).ToShortDateString());       //date
                    objNewRecords.Add(obj);

                    string rangeToUpdate = "DailyCheckAll!A" + counter + ":A" + counter;

                    SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                    request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    var response = request.Execute();



                    ////Employee
                    objNewRecords = new List<IList<Object>>();
                    obj = new List<Object>();
                    obj.Add(biz);
                    obj.Add(x.EmpName);
                    obj.Add(x.EmpStartTime);
                    objNewRecords.Add(obj);

                    rangeToUpdate = "DailyCheckAll!C" + counter + ":E" + counter;

                    request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                    request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    response = request.Execute();


                    ////Imported
                    objNewRecords = new List<IList<Object>>();
                    obj = new List<Object>();
                    obj.Add(DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                    objNewRecords.Add(obj);

                    rangeToUpdate = "DailyCheckAll!S" + counter + ":S" + counter;

                    request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                    request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                    request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    response = request.Execute();


                    counter = counter + 1;

                }


            }


        }


        public void UpdateShackSquareTransactions(string sheetID, List<SquareDataShack> localList)
        {
            Utilities.ConfigSettings cs = new Utilities.ConfigSettings();

            //** need to get last row of data
            string worksheetName = cs.ReturnConfigSetting("NightlyRouteToSlack", "ShackRouteJournalLastRow").ToString();
            worksheetName = worksheetName.Substring(0, worksheetName.IndexOf("!"));
            int counter = ReturnLastRowOfDataShack(sheetID);

            foreach (SquareDataShack x in localList)
            {

                //Initial Job Data
                List<IList<Object>> objNewRecords = new List<IList<Object>>();
                IList<Object> obj = new List<Object>();
                obj.Add(x.RouteDate);       //date
                obj.Add(x.TotalCollected);           //total collected
                obj.Add(x.PaymentType);           //payment type
                obj.Add(x.PanSuffix);         //PAN suffix
                obj.Add(x.Tip);             //tip
                obj.Add(x.Tax);             //tax
                objNewRecords.Add(obj);

                string rangeToUpdate = worksheetName + "!E" + counter + ":J" + counter;

                SpreadsheetsResource.ValuesResource.AppendRequest request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var response = request.Execute();



                //Opp ID
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add(x.JobID);
                objNewRecords.Add(obj);

                rangeToUpdate = worksheetName + "!R" + counter + ":R" + counter;

                request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();



                //Imported
                objNewRecords = new List<IList<Object>>();
                obj = new List<Object>();
                obj.Add(DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                objNewRecords.Add(obj);

                rangeToUpdate = worksheetName + "!U" + counter + ":U" + counter;

                request = _sheetService.Spreadsheets.Values.Append(new ValueRange() { Values = objNewRecords }, sheetID, rangeToUpdate);
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                response = request.Execute();





                counter = counter + 1;

            }


        }


    }
}
