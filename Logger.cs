using SeleniumFrameWork.Base;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.Helpers
{
    public class Logger:HookBase
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



        public static void log(string msg)
        {
            
          

            StackTrace stackTrace = new StackTrace();
            // get calling method name
            String testName = stackTrace.GetFrame(1).GetMethod().Name;
            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);
            String methodName = " [ " + testName + " ] ";
            sw.WriteLine(sLogFormat + msg);

            sw.Flush();

            sw.Close();

            

        }

        public static string logPath()
        {
            string asmblyPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actPath = asmblyPath.Substring(0, asmblyPath.LastIndexOf("bin"));
            string sltnPath = new Uri(actPath).LocalPath;
            string logPath = sltnPath + @"TestLogs\";
            return logPath;
        }



        public static void testName(string testCaseName)

        {



            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);

            sw.WriteLine(sLogFormat + testCaseName);

            sw.Flush();

            sw.Close();

        }

        //Assert
        public static void assertTest(string assertion)

        {
            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);

            sw.WriteLine(sLogFormat + assertion);

            sw.Flush();

            sw.Close();

        }
        //Debug
        public static void debugTest(string debugmsg)

        {
            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);

            sw.WriteLine(sLogFormat + debugmsg);

            sw.Flush();

            sw.Close();

        }
        //Write
        public static void writeTestSteps(string writemsg)

        {
            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);

            sw.WriteLine(sLogFormat + writemsg);

            sw.Flush();

            sw.Close();

        }
        public static void cleanbtwTestRuns(string msg)

        {
            StreamWriter sw = new StreamWriter(logPath() + timeStamp() + @"_log.txt", true);

            sw.WriteLine(msg);

            sw.Flush();

            sw.Close();

        }




    }
}

