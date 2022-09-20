using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace NightlyRouteToSlack.Utilities
{
    public class ConfigSettings
    {

        public string ReturnConfigSetting(string appName, string settingKey)
        {
            string settingValue = "";

            dbConnect dc = new dbConnect();
            dc.OpenSettingsConnection();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = dc.conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetKeyValuePairs";

            cmd.Parameters.Add("@appName", SqlDbType.NVarChar);
            cmd.Parameters["@appName"].Value = appName;

            cmd.Parameters.Add("@settingKey", SqlDbType.NVarChar);
            cmd.Parameters["@settingKey"].Value = settingKey;

            DataTable dt = new DataTable();

            SqlDataAdapter ds = new SqlDataAdapter(cmd);
            ds.Fill(dt);
            dc.CloseConnection();

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["settingValue"].ToString()))
                {
                    settingValue = dt.Rows[0]["settingValue"].ToString();
                }
            }


            return settingValue;
        }


    }
}
