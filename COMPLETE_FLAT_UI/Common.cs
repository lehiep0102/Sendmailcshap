using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace COMPLETE_FLAT_UI
{
    class Common
    {   
        public static string ORACLE_SERVER;
        public static string ORACLE_PORT;
        public static string ORACLE_SERVICE_NAME;
        public static string ORACLE_USER;
        public static string ORACLE_PASSWORD;
        public static string SMTP_SERVER;
        public static string SMTP_PORT;
        public static string MAIL_FROM;
        public static string MAIL_FROM_NAME;
        public static string SMTP_USERNAME;
        public static string SMTP_PASSWORD;
        public static string EmailCC;
        public static string EmailBCC;
        public static string GetSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
    }
}
