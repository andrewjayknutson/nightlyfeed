using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace NightlyRouteToSlack
{
    class dbConnect
    {
        public SqlConnection conn { get; set; }

        public void OpenConnection()
        {
            string conString = ConfigurationManager.AppSettings["Consumer_DSN"].ToString().Replace("\\\\","\\");
            conn = new SqlConnection(conString);
            conn.Open();
        }

        public void OpenMessageConnection()
        {
            string conString = ConfigurationManager.AppSettings["Message_DSN"].ToString().Replace("\\\\", "\\");
            conn = new SqlConnection(conString);
            conn.Open();
        }

        public void OpenSettingsConnection()
        {
            string conString = ConfigurationManager.AppSettings["Settings_DSN"].ToString().Replace("\\\\", "\\");
            conn = new SqlConnection(conString);
            conn.Open();
        }

        public void OpenArchiveConnection()
        {
            string conString = ConfigurationManager.AppSettings["Archive_DSN"].ToString().Replace("\\\\", "\\");
            conn = new SqlConnection(conString);
            conn.Open();
        }

        public void CloseConnection()
        {
            conn.Close();
        }

    }
}
