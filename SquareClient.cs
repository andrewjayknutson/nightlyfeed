using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Linq;
using Newtonsoft.Json;
using System.Globalization;

namespace Utilities360Wow
{
    public class SquareClient
    {
        //***********************
        //*** Using Connect v2 **
        //***********************
        public List<JObject> RetrieveSquarePaymentsReceived(Boolean getAll, string accessToken)
        {

            //TODO ... add a try catch and if something errors, send back a message letting me know it errored on retrieving payments from Square

            SquareData local = new SquareData();
            List<SquareData> localList = new List<SquareData>();

            //TESTING
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/payments?begin_time=" + DateTime.Now.AddDays(-5).ToString("yyyy-MM-dd") + "T00:00:00-06:00&end_time=" + DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + "T23:59:59-06:00");

            //PRODUCTION
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/payments?begin_time=" + DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd") + "T00:00:00-06:00&end_time=" + DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + "T23:59:59-06:00");

            webReq.UseDefaultCredentials = true;
            webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //remove extraneous beginning and ending data
            contents = contents.Substring(13, (contents.Length - 15));

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            return obj;

        }







        //**************************
        //*** Using Square API v1 **
        //**************************

        public List<SquareTransactions> RetrieveSquareDepositTransactions(string depositStartDate, string depositEndDate, string locationID, string accessToken)
        {
            //local variables
            DateTime selectedStartDate = DateTime.Parse(depositStartDate).AddDays(-1);
            DateTime selectedEndDate = DateTime.Parse(depositEndDate);
            SquareTransactions local = new SquareTransactions();
            List<SquareTransactions> localList = new List<SquareTransactions>();

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + selectedStartDate.ToString("yyyy-MM-dd") + "T19:00:00-0600&end_time=" + selectedEndDate.ToString("yyyy-MM-dd") + "T18:59:59-0600");
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                foreach (JToken tokens in x.SelectToken("tender"))
                {

                    if (tokens["type"].ToString().ToLower() == "credit_card")
                    {
                        local = new SquareTransactions();

                        local.TransID = x["id"].ToString();
                        local.TransDate = DateTime.Parse(x["created_at"].ToString()).AddHours(-6).ToString();
                        local.TransCollected = (float.Parse(tokens["total_money"]["amount"].ToString()) / 100).ToString();
                        local.TransGrossSales = (float.Parse(x["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                        local.TransTax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                        local.TransTip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();
                        local.TransType = tokens["type"].ToString();
                        local.TransFees = (float.Parse(x["processing_fee_money"]["amount"].ToString()) / 100).ToString();

                        //add the processing fee as it is a negative number
                        local.TransNet = ((float.Parse(tokens["total_money"]["amount"].ToString()) + float.Parse(x["processing_fee_money"]["amount"].ToString())) / 100).ToString();

                        localList.Add(local);
                    }

                }


                foreach (JToken tokens in x.SelectToken("refunds"))
                {

                    local = new SquareTransactions();

                    local.TransID = x["id"].ToString();
                    local.TransDate = DateTime.Parse(x["created_at"].ToString()).AddHours(-6).ToString();
                    local.TransCollected = (float.Parse(tokens["refunded_money"]["amount"].ToString()) / 100).ToString();
                    local.TransGrossSales = "0.00";
                    local.TransTax = "0.00";
                    local.TransTip = "0.00";
                    local.TransType = tokens["type"].ToString();
                    local.TransFees = (float.Parse(tokens["refunded_processing_fee_money"]["amount"].ToString()) / 100).ToString();

                    //add the processing fee as it is a negative number
                    local.TransNet = ((float.Parse(tokens["refunded_money"]["amount"].ToString()) + float.Parse(tokens["refunded_processing_fee_money"]["amount"].ToString())) / 100).ToString();

                    localList.Add(local);

                }



            }


            return localList;


        }

        public List<SquareDeposit> RetrieveSquareDeposits(string depositStartDate, string depositEndDate, string locationID, string accessToken)
        {
            //local variables
            DateTime selectedStartDate = DateTime.Parse(depositStartDate);
            DateTime selectedEndDate = DateTime.Parse(depositEndDate);
            SquareDeposit local = new SquareDeposit();
            List<SquareDeposit> localList = new List<SquareDeposit>();

            //go out and grab the data from Square
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/settlements?begin_time=" + selectedStartDate.ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + selectedEndDate.ToString("yyyy-MM-dd") + "T23:59:59-0600");
            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                local = new SquareDeposit();

                local.DepositID = x["id"].ToString();
                local.DepositDate = DateTime.Parse(x["initiated_at"].ToString()).AddHours(-6).ToString();
                local.DepositAmount = (float.Parse(x["total_money"]["amount"].ToString()) / 100).ToString();

                localList.Add(local);

            }


            return localList;

        }


        public List<SquareDataMove> RetrieveSquareTransactionsMove(Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            SquareDataMove local = new SquareDataMove();
            List<SquareDataMove> localList = new List<SquareDataMove>();

            //TESTING - go out and grab the data from Square
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T23:59:59-0600");

            //PRODUCTION
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + "T23:59:59-0600");

            webReq.UseDefaultCredentials = true; 
            webReq.Headers.Add("Authorization", "Bearer " + accessToken);

            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                local = new SquareDataMove();

                local.RouteDate = x["created_at"].ToString().Substring(0, x["created_at"].ToString().IndexOf("/202") + 5);
                local.TotalCollected = (float.Parse(x["total_collected_money"]["amount"].ToString()) / 100).ToString();
                local.Tax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                local.Tip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();
                local.TaxableRevenue = "0";
                local.Total = (float.Parse(local.TotalCollected) - float.Parse(local.Tip) - float.Parse(local.Tax)).ToString();

                foreach (JToken tender in x.SelectToken("tender"))
                {
                    if (tender["type"].ToString().ToLower() == "credit_card")
                    {
                        local.PaymentType = tender["card_brand"].ToString().Replace("MASTER_CARD", "Mastercard").Replace("AMERICAN_EXPRESS", "AMEX");
                        local.PanSuffix = tender["pan_suffix"].ToString();
                        local.SwipeOrKeyed = tender["entry_method"].ToString();
                    }
                }

                foreach (JToken tender in x.SelectToken("itemizations"))
                {
                    if (tender["name"].ToString().ToLower() == "custom amount")
                    {
                        try
                        {
                            local.JobID = tender["notes"].ToString();
                        }
                        catch
                        {

                        }
                    }
                }

                //I want to get them all
                if (getAll)
                {
                    localList.Add(local);
                }

            }


            return localList;

        }



        public List<SquareDataShack> RetrieveSquareTransactionsShack(Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            SquareDataShack local = new SquareDataShack();
            List<SquareDataShack> localList = new List<SquareDataShack>();

            //go out and grab the data from Square
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + "T23:59:59-0600");

            //used for testing
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-5).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T23:59:59-0600");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                local = new SquareDataShack();

                local.RouteDate = x["created_at"].ToString().Substring(0, x["created_at"].ToString().IndexOf("/202") + 5);
                local.TotalCollected = (float.Parse(x["total_collected_money"]["amount"].ToString()) / 100).ToString();
                local.Tax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                local.Tip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();
                local.Total = (float.Parse(local.TotalCollected) - float.Parse(local.Tip) - float.Parse(local.Tax)).ToString();

                foreach (JToken tender in x.SelectToken("tender"))
                {
                    if (tender["type"].ToString().ToLower() == "credit_card")
                    {
                        local.PaymentType = tender["card_brand"].ToString().Replace("MASTER_CARD", "MasterCard").Replace("AMERICAN_EXPRESS", "AMEX");
                        local.PanSuffix = tender["pan_suffix"].ToString();
                    }

                }

                foreach (JToken itemization in x.SelectToken("itemizations"))
                {
                    try
                    {
                        local.JobID = itemization["notes"].ToString();
                    }
                    catch
                    {

                    }

                    if (itemization["name"].ToString().ToLower().Trim() == "tax")
                    {
                        local.Tax = (float.Parse(itemization["total_money"]["amount"].ToString()) / 100).ToString();
                    }


                }

                //I want to get them all
                if (getAll && float.Parse(local.Total) > 0)
                {
                    localList.Add(local);
                }

            }


            return localList;

        }




        public List<SquareSwiped> RetrieveSquareSwipedDates(Boolean getAll, string locationID, string accessToken, string startDate, string endDate)
        {
            //local variables
            string cursor = "";
            SquareSwiped local = new SquareSwiped();
            List<SquareSwiped> localList = new List<SquareSwiped>();

            //a whole bunch of days ago
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Now.AddDays(-21).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59Z");

            //yesterday
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59Z");


            while (cursor != "end")
            {
                //specific dates
                HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Parse(startDate).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Parse(endDate).ToString("yyyy-MM-dd") + "T23:59:59Z" + cursor);

                webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
                webReq.Method = "GET";
                HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
                Stream answer = webResp.GetResponseStream();
                StreamReader _answer = new StreamReader(answer);
                string contents = _answer.ReadToEnd();

                //convert the data to json
                contents = "[" + contents + "]";
                var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

                //read through each jbson object
                foreach (JObject x in obj)
                {
                    try
                    {
                        cursor = "&cursor=" + x["cursor"].ToString();
                    }
                    catch
                    {
                        cursor = "end";
                    }
                  
                    foreach (JToken transactions in x.SelectToken("transactions"))
                    {
                        foreach (JToken tenders in transactions.SelectToken("tenders"))
                        {
                            if (tenders["type"].ToString().ToLower() == "card")
                            {
                                local = new SquareSwiped();

                                local.Revenue = (float.Parse(tenders["amount_money"]["amount"].ToString()) / 100).ToString();
                                local.ProceesingFee = (float.Parse(tenders["processing_fee_money"]["amount"].ToString()) / 100).ToString();
                                local.ProcessingType = tenders["card_details"]["entry_method"].ToString().ToLower();

                                localList.Add(local);
                            }
                        }
                    }
                }
            }




            return localList;

        }



        public List<SquareSwiped> RetrieveSquareSwipedByDates(DateTime startDate, DateTime endDate, Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            SquareSwiped local = new SquareSwiped();
            List<SquareSwiped> localList = new List<SquareSwiped>();

            //a whole bunch of days ago
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Now.AddDays(-21).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59Z");

            //yesterday
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + startDate.ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + endDate.ToString("yyyy-MM-dd") + "T23:59:59Z");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            if (!(contents == "{}"))
            {

                //convert the data to json
                contents = "[" + contents + "]";
                var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

                //read through each jbson object
                foreach (JObject x in obj)
                {
                    foreach (JToken transactions in x.SelectToken("transactions"))
                    {
                        foreach (JToken tenders in transactions.SelectToken("tenders"))
                        {
                            if (tenders["type"].ToString().ToLower() == "card")
                            {
                                local = new SquareSwiped();

                                local.Revenue = (float.Parse(tenders["amount_money"]["amount"].ToString()) / 100).ToString();
                                local.ProceesingFee = (float.Parse(tenders["processing_fee_money"]["amount"].ToString()) / 100).ToString();
                                local.ProcessingType = tenders["card_details"]["entry_method"].ToString().ToLower();

                                localList.Add(local);
                            }
                        }
                    }
                }

            }


            return localList;

        }


        public List<SquareSwiped> RetrieveSquareSwiped(Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            SquareSwiped local = new SquareSwiped();
            List<SquareSwiped> localList = new List<SquareSwiped>();

            //a whole bunch of days ago
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Now.AddDays(-21).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59Z");

            //yesterday
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v2/locations/" + locationID + "/transactions?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00Z&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59Z");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            contents = "[" + contents + "]";
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                foreach (JToken transactions in x.SelectToken("transactions"))
                {
                    foreach (JToken tenders in transactions.SelectToken("tenders"))
                    {
                        if (tenders["type"].ToString().ToLower() == "card")
                        {
                            local = new SquareSwiped();

                            local.Revenue = (float.Parse(tenders["amount_money"]["amount"].ToString()) / 100).ToString();
                            local.ProceesingFee = (float.Parse(tenders["processing_fee_money"]["amount"].ToString()) / 100).ToString();
                            local.ProcessingType = tenders["card_details"]["entry_method"].ToString().ToLower();

                            localList.Add(local);
                        }
                    }
                }
            }

            return localList;

        }



        public List<SquareData> RetrieveSquareTransactions(Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            float revenue = 0;
            bool jobIdFound = false;
            SquareData local = new SquareData();
            List<SquareData> localList = new List<SquareData>();

            //used for testing ...
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-4).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59-0600");
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T23:59:59-0600");
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + "T23:59:59-0600");

            //set back to this for production
            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59-0600");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach(JObject x in obj)
            {
                local = new SquareData();

                local.JobID = "";
                local.Route = "";
                local.TVLargeCount = "0";
                local.TVLargeValue = "0";
                local.TVSmallCount= "0";
                local.TVSmallValue = "0";
                local.TiresCount = "0";
                local.TiresValue = "0";
                local.MattressCount = "0";
                local.MattressValue = "0";
                local.JobType = "Residential";
                local.PayMethod = "Credit No Swipe";

                local.RouteDate = x["created_at"].ToString().Substring(0, x["created_at"].ToString().IndexOf("/202") + 5);
                local.Discount = (float.Parse(x["discount_money"]["amount"].ToString()) / 100).ToString();
                local.Tax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                local.Tip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();

                local.Revenue = "0";    //need to figure a way to sum all "load", "min", "standard" items
                local.NetRevenue = (float.Parse(x["net_sales_money"]["amount"].ToString()) / 100).ToString();         //includes +"extras" -"discounts"

                //check for credit swipe vs. not swipe
                foreach (JToken tender in x.SelectToken("tender"))
                {
                    if (tender["type"].ToString().ToLower() == "credit_card")
                    {
                        if (tender["entry_method"].ToString().ToLower() == "swiped")
                        {
                            local.PayMethod = "Credit Swipe";
                        }
                    }
                }


                foreach (JToken itemizations in x.SelectToken("itemizations"))
                {

                    //get Job ID if "load" or "standard item price" or "min junk removal" item
                    if ((itemizations["name"].ToString().ToLower().Contains("load") || itemizations["name"].ToString().ToLower().Contains("standard item price") || itemizations["name"].ToString().ToLower().Contains("min junk removal") || itemizations["name"].ToString().ToLower().Contains("custom amount")) & !jobIdFound)
                    {
                        try
                        {
                            local.JobID = itemizations["notes"].ToString();
                            jobIdFound = true;
                        }
                        catch
                        {

                        }
                    }

                    //"truck" item
                    if (itemizations["name"].ToString().ToLower().Contains("truck"))
                    {
                        local.Route = itemizations["name"].ToString();
                    }

                    //"residential" vs. "commercial" item
                    if (itemizations["name"].ToString().ToLower().Contains("commercial"))
                    {
                        local.JobType = "Commercial";
                    }

                    //calculate Revenue if "load" or "standard item price" or "min junk removal" item
                    if (itemizations["name"].ToString().ToLower().Contains("load") || itemizations["name"].ToString().ToLower().Contains("standard item price") || itemizations["name"].ToString().ToLower().Contains("min junk removal") || itemizations["name"].ToString().ToLower().Contains("custom amount"))
                    {
                        revenue = revenue + (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100);
                    }


                    //"mattress" item
                    if (itemizations["name"].ToString().ToLower().Contains("mattress"))
                    {
                        local.MattressCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.MattressValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"tires" item
                    if (itemizations["name"].ToString().ToLower().Contains("tire charge"))
                    {
                        local.TiresCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TiresValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"large tv" item
                    if (itemizations["name"].ToString().ToLower().Contains("large tv"))
                    {
                        local.TVLargeCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TVLargeValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"small tv" item
                    if (itemizations["name"].ToString().ToLower().Contains("small tv"))
                    {
                        local.TVSmallCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TVSmallValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                }

                local.Revenue = revenue.ToString();

                revenue = 0;
                jobIdFound = false;

                //I want to get them all
                if (getAll)
                {
                    localList.Add(local);
                }

                //I want to only get the good ones 
                if (!getAll & !string.IsNullOrEmpty(local.JobID) & !string.IsNullOrEmpty(local.Route))
                {
                    localList.Add(local);
                }


                //REMOVE BEFORE LAUNCHING
                //if (!string.IsNullOrEmpty(local.Route))
                //{
                //    localList.Add(local);
                //}


            }


            return localList;

        }




        public List<SquareData> RetrieveOvernightSquareTransactions(Boolean getAll, string locationID, string accessToken)
        {
            //local variables
            float revenue = 0;
            bool jobIdFound = false;
            SquareData local = new SquareData();
            List<SquareData> localList = new List<SquareData>();

            //used for testing because it pulls everything from the previous day vs. at 11:50 at night
            //HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T13:00:00-0600&end_time=" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "T15:59:59-0600");

            HttpWebRequest webReq = (HttpWebRequest)WebRequest.Create("https://connect.squareup.com/v1/" + locationID + "/payments?begin_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T00:00:00-0600&end_time=" + DateTime.Now.ToString("yyyy-MM-dd") + "T23:59:59-0600");

            webReq.UseDefaultCredentials = true; webReq.Headers.Add("Authorization", "Bearer " + accessToken);
            webReq.Method = "GET";
            HttpWebResponse webResp = (HttpWebResponse)webReq.GetResponse();
            Stream answer = webResp.GetResponseStream();
            StreamReader _answer = new StreamReader(answer);
            string contents = _answer.ReadToEnd();

            //convert the data to json
            var obj = JsonConvert.DeserializeObject<JArray>(contents).ToObject<List<JObject>>();

            //read through each jbson object
            foreach (JObject x in obj)
            {
                local = new SquareData();

                local.JobID = "";
                local.Route = "";
                local.TVLargeCount = "0";
                local.TVLargeValue = "0";
                local.TVSmallCount = "0";
                local.TVSmallValue = "0";
                local.TiresCount = "0";
                local.TiresValue = "0";
                local.MattressCount = "0";
                local.MattressValue = "0";
                local.JobType = "Residential";
                local.PayMethod = "Credit No Swipe";

                local.RouteDate = x["created_at"].ToString().Substring(0, x["created_at"].ToString().IndexOf(" "));
                local.Discount = (float.Parse(x["discount_money"]["amount"].ToString()) / 100).ToString();
                local.Tax = (float.Parse(x["tax_money"]["amount"].ToString()) / 100).ToString();
                local.Tip = (float.Parse(x["tip_money"]["amount"].ToString()) / 100).ToString();
                local.RMTip = "0";
                local.RMLink = "";
                local.RMMatches = "FALSE";

                local.Revenue = "0";    //need to figure a way to sum all "load", "min", "standard" items
                local.NetRevenue = (float.Parse(x["net_sales_money"]["amount"].ToString()) / 100).ToString();         //includes +"extras" -"discounts"

                //check for credit swipe vs. not swipe
                foreach (JToken tender in x.SelectToken("tender"))
                {
                    if (tender["type"].ToString().ToLower() == "credit_card")
                    {
                        if (tender["entry_method"].ToString().ToLower() == "swiped")
                        {
                            local.PayMethod = "Credit Swipe";
                        }
                    }
                }


                foreach (JToken itemizations in x.SelectToken("itemizations"))
                {

                    //get Job ID if "load" or "standard item price" or "min junk removal" item
                    if ((itemizations["name"].ToString().ToLower().Contains("load") || itemizations["name"].ToString().ToLower().Contains("standard item price") || itemizations["name"].ToString().ToLower().Contains("min junk removal") || itemizations["name"].ToString().ToLower().Contains("custom amount") || itemizations["name"].ToString().ToLower().Contains("full truck")) & !jobIdFound)
                    {
                        try
                        {
                            local.JobID = itemizations["notes"].ToString();
                            jobIdFound = true;
                        }
                        catch
                        {

                        }
                    }

                    //"truck" item
                    if (itemizations["name"].ToString().ToLower().Contains("truck"))
                    {
                        local.Route = itemizations["name"].ToString();
                    }

                    //"residential" vs. "commercial" item
                    if (itemizations["name"].ToString().ToLower().Contains("commercial"))
                    {
                        local.JobType = "Commercial";
                    }

                    //calculate Revenue if "load" or "standard item price" or "min junk removal" item
                    if (itemizations["name"].ToString().ToLower().Contains("load") || itemizations["name"].ToString().ToLower().Contains("standard item price") || itemizations["name"].ToString().ToLower().Contains("min junk removal") || itemizations["name"].ToString().ToLower().Contains("custom amount") || itemizations["name"].ToString().ToLower().Contains("full truck"))
                    {
                        revenue = revenue + (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100);
                    }


                    //"mattress" item
                    if (itemizations["name"].ToString().ToLower().Contains("mattress"))
                    {
                        local.MattressCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.MattressValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"tires" item
                    if (itemizations["name"].ToString().ToLower().Contains("tire charge"))
                    {
                        local.TiresCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TiresValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"large tv" item
                    if (itemizations["name"].ToString().ToLower().Contains("large tv"))
                    {
                        local.TVLargeCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TVLargeValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                    //"small tv" item
                    if (itemizations["name"].ToString().ToLower().Contains("small tv"))
                    {
                        local.TVSmallCount = (float.Parse(itemizations["quantity"].ToString()) / 1).ToString();
                        local.TVSmallValue = (float.Parse(itemizations["gross_sales_money"]["amount"].ToString()) / 100).ToString();
                    }

                }

                local.Revenue = revenue.ToString();

                revenue = 0;
                jobIdFound = false;

                //I want to get them all
                if (getAll)
                {
                    localList.Add(local);
                }

                //I want to only get the good ones 
                if (!getAll & !string.IsNullOrEmpty(local.JobID) & !string.IsNullOrEmpty(local.Route))
                {
                    localList.Add(local);
                }


                //REMOVE BEFORE LAUNCHING
                //if (!string.IsNullOrEmpty(local.Route))
                //{
                //    localList.Add(local);
                //}


            }


            return localList;

        }














    }

}
