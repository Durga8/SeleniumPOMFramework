using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SeleniumFrameWork.Base.DriverSetUp;

namespace SeleniumFrameWork.Config
{
    public class Settings
    {
        public static int Timeout { get; set; }

        public static string IsReporting { get; set; }
        public static string Browser { get; set; }
        public static string TestType { get; set; }

        public static string Name { get; set; }

        public static string AUT { get; set; }

        public static string BuildName { get; set; }
        public static SqlConnection ApplicationCon { get; set; }

        public static OracleConnection AppConn { get; set; }
        public static string AppConnectionString { get; set; }
        //public static string IsLog { get; set; }

        //public static string logPath { get; set; }
        private static bool _fileCreated = false;
        public static bool FileCreated
        {
            get
            {
                return _fileCreated;
            }
            set
            {
                _fileCreated = value;
            }
        }

    }
}
