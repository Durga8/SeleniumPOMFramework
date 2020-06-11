using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.Helpers
{
    public static class DriverContext
    {
        public static string getTestCaseName()
        {
            string currentTest = NUnit.Framework.TestContext.CurrentContext.Test.Name;
            
            return currentTest;
        }

        public static string getSolutionPath()
        {
            string asmblyPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actPath = asmblyPath.Substring(0, asmblyPath.LastIndexOf("bin"));
            string sltnPath = new Uri(actPath).LocalPath;
            return sltnPath;
        }
    }
}
