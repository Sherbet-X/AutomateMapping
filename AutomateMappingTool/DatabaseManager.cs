using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Oracle.ManagedDataAccess.Client;

namespace AutomateMappingTool
{
    class DatabaseManager
    {
        public static OracleConnection ConnectionProd { get; private set; }
        public static OracleConnection ConnectionTemp { get; private set; }

        private string user;
        private string password;


        public string Connection_User { set { user = value; } }
        public string Connection_Password { set { password = value; } }

        private void CreateConnectionProd()
        {

        }

        private void CreateConnectionTemp()
        {

        }

    }
}
