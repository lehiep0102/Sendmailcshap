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

            string host = "10.0.25.24";
            int port = 1521;
            string sid = "oracledb";
            string user = "quote";
            string password = "123456";
            //
            /* //Load config file
             string logfile = System.IO.Directory.GetCurrentDirectory();
             string datastring = "";
             byte[] data = new byte[2049];
             byte[] decrypt = new byte[2049];
             int lenbyte = 0;
             DataCryption crypt = new DataCryption();

             string host = "";
             int port = 1521;
             string sid = "";
             string user = "";
             string password = "";

             try
             {
                 FileStream fs = new FileStream(logfile + "\\data\\config.cfg", FileMode.OpenOrCreate, FileAccess.Read);
                 fs.Read(data, 0, 2048);
                 fs.Close();

                 for (int i = 0; i < data.Length; i++)
                 {
                     if (data[i] == (byte)0)
                         break;
                     lenbyte++;
                 }
                 //  if (lenbyte <= 0)
                 //     return;

                 byte[] fixdata = new byte[lenbyte];

                 Buffer.BlockCopy(data, 0, fixdata, 0, lenbyte);

                 string aaaa = Encoding.UTF8.GetString(fixdata);

                 string cccc = aaaa.Substring(1, aaaa.Length - 1);

                 datastring = Encoding.UTF8.GetString(crypt.decrypt(Encoding.UTF8.GetBytes(cccc)));

                 string[] split = datastring.Split('|');

                 if (split.Length > 3)
                 {
                     host = split[0];
                     sid = split[1];
                     user = split[2];
                     password = split[3];
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }
             //MessageBox.Show(host, "aaaa"); 
             //MessageBox.Show(sid, "aaaa"); 
             //MessageBox.Show(user, "aaaa"); 
             //MessageBox.Show(password,"aaaa");
             */

            return DBOracleUtils.GetDBConnection(host, port, sid, user, password.Trim());
        }
    }
}
