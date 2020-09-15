using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;

namespace MAS_EMAIL
{
    class DBUtilsreal
    {
        public static OracleConnection GetDBConnection()
        {

            string host = "10.0.31.15";
            int port = 1521;
            string sid = "oracledb";
            string user = "quote";
            string password = "123456";
            //

            //MessageBox.Show(host, "aaaa"); 
            //MessageBox.Show(sid, "aaaa"); 
            //MessageBox.Show(user, "aaaa"); 
            //MessageBox.Show(password,"aaaa"); 

            return DBOracleUtilsreal.GetDBConnectionreal(host, port, sid, user, password.Trim());
        }
    }
    class DBOracleUtilsreal
    {
        public static OracleConnection
                   GetDBConnectionreal(string host, int port, String sid, String user, String password)
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
