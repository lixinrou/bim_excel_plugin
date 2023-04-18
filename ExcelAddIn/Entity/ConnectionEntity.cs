using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Entity
{
    public class ConnectionEntity
    {
        public string ConnectionName { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Host { get; set; } 
        public string Database { get; set; }
        public int ConnectionTimeout { get; set; } = 500;
        public string Provider { get; set; } = "MySql.Data.MySqlClient";
    }
}
