using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Common
{
    public class DataSource
    {
        private string appsetting = string.Empty;

        public DataSource()
        {
            Settings settings = new Settings();
            string current = settings.GetSettingValue("active");
            appsetting = ConfigurationManager.ConnectionStrings[current].ToString();
        }

        public DbConnection GetCurrentConnection(string type)
        {
            if (type.Equals("mysql"))
            {
                return new MySqlConnection(appsetting);
            }
            else if (type.Equals("sqlserver"))
            {
                return new SqlConnection(appsetting);
            }
            else
            {
                return null;
            }
        }
    }
}
