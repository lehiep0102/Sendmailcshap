using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;

namespace COMPLETE_FLAT_UI
{
    class DBUtils
    {
        public static OracleConnection GetDBConnection()
        {

            string host = Common.ORACLE_SERVER;
            int port = Int32.Parse(Common.ORACLE_PORT);
            string sid = Common.ORACLE_SERVICE_NAME; 
            string user = Common.ORACLE_USER;
            string password = Common.ORACLE_PASSWORD;
  
            return DBOracleUtils.GetDBConnection(host, port, sid, user, password.Trim());
        }
    }
}
