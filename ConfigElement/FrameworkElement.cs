using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.ConfigElement
{
    public class FrameworkElement:ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name { get { return (string)base["name"]; } }

        [ConfigurationProperty("aut", IsRequired = true)]
        public string AUT { get { return (string)base["aut"]; } }

        [ConfigurationProperty("browser", IsRequired = true)]
        public string Browser { get { return (string)base["browser"]; } }

        [ConfigurationProperty("testType", IsRequired = false)]
        public string TestType { get { return (string)base["testType"]; } }

        [ConfigurationProperty("isLog", IsRequired = false)]
        public string IsLog { get { return (string)base["isLog"]; } }

        [ConfigurationProperty("logPath", IsRequired = false)]
        public string LogPath { get { return (string)base["logPath"]; } }

        [ConfigurationProperty("applicationConn", IsRequired = false)]
        public SqlConnection ApplicationCon { get { return (SqlConnection)base["applicationConn"]; } }

        [ConfigurationProperty("appConn", IsRequired = false)]
        public OracleConnection AppConn { get { return (OracleConnection)base["appConn"]; } }

        [ConfigurationProperty("applicationDB", IsRequired = false)]
        public string AppConnectionString { get { return (string)base["applicationDB"]; } }


    }
}
