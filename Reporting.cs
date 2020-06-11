using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SeleniumFrameWork.Helpers
{
    public class Reporting
    {
        private static string sLogFormat;

        private static string sErrorTime;


        public static string timeStamp()

        {

            //sLogFormat used to create log files format :



            sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";



            //this variable used to create log filename format "



            string sYear = DateTime.Now.Year.ToString();

            string sMonth = DateTime.Now.Month.ToString();

            string sDay = DateTime.Now.Day.ToString();

            sErrorTime = sYear + sMonth + sDay;



            return sErrorTime;

        }
        


    }
}

