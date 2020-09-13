using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;

namespace COMPLETE_FLAT_UI
{
    class DBOracleUtils
    {
        public static OracleConnection
                  GetDBConnection(string host, int port, String sid, String user, String password)
        {

            // Connection String để kết nối trực tiếp tới Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
                 + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SID="//(SERVICE_NAME = 
                 + sid + ")));Password=" + password + ";User ID=" + user;

            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = connString;
            return conn;
        }

        internal static OracleConnection GetDBConnection()
        {
            throw new NotImplementedException();
        }
    }
}
