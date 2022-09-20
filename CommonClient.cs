using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace Utilities360Wow
{
    public class CommonClient
    {

        public string GetEOMDate()
        {
            DateTime today = DateTime.Today;
            DateTime endOfMonth;

            if (DateTime.Today.Day.ToString() == "1")
            {
                endOfMonth = new DateTime(today.Year, today.Month, 1).AddDays(-1);
            }
            else 
            {
                endOfMonth = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));
            }

            return endOfMonth.ToString();
        }


        public string GetStartDate()
        {
            string localMonth = DateTime.Today.Month.ToString();
            string localYear = DateTime.Today.Year.ToString();

            if (DateTime.Today.Day.ToString() == "1")
            {
                localMonth = DateTime.Today.AddDays(-1).Month.ToString();
            }

            if (DateTime.Today.Month.ToString() == "1")
            {
                localYear = DateTime.Today.AddDays(-1).Year.ToString();
            }

            return localMonth + "/1/" + localYear;
        }

        public double CalculateDailyRevenueNeed(double totalRev, double goalRev, DateTime start, DateTime end)
        {
            double pace = 0;
            double remainderRev = 0;
            double numberOfDays = 0;

            numberOfDays = (end - start).TotalDays + 1;

            remainderRev = goalRev - totalRev;

            pace = remainderRev / numberOfDays;

            switch (DateTime.Today.AddDays(-1).Month)
            {
                case 1:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 2:
                    pace = remainderRev / (28 - numberOfDays);
                    break;
                case 3:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 4:
                    pace = remainderRev / (30 - numberOfDays);
                    break;
                case 5:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 6:
                    pace = remainderRev / (30 - numberOfDays);
                    break;
                case 7:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 8:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 9:
                    pace = remainderRev / (30 - numberOfDays);
                    break;
                case 10:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
                case 11:
                    pace = remainderRev / (30 - numberOfDays);
                    break;
                case 12:
                    pace = remainderRev / (31 - numberOfDays);
                    break;
            }

            return pace;

        }

       

    }
}
