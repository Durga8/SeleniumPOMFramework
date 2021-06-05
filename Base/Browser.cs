using NUnit.Framework;
using SeleniumFrameWork.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SeleniumFrameWork.Base
{
    public class Browser:HookBase
    {
        

        public static void changeBrowser()
        {
           
            string crosBrwsr = getParameter("CrossBrowser");
            int iter = getIteration();
            if (crosBrwsr.Equals("Yes"))
            {
                openExcelRunConfig("RunConfigurationManager", "TestScripts");
                string testcaseID = DriverContext.getTestCaseName();
                int rowNumber = xlWorkSheet.Columns.Find(testcaseID).Cells.Row;
                int columnNumber = xlWorkSheet.Columns.Find("Browser").Cells.Column;
                string brswr = xlWorkSheet.Cells[rowNumber, columnNumber].Text.ToString();

                if (brswr.Equals("Chrome"))
                {
                    xlWorkSheet.Cells[rowNumber, columnNumber] = "IE";
                   

                }
                else if (brswr.Equals("IE"))
                {
                    xlWorkSheet.Cells[rowNumber, columnNumber] = "Chrome";
                }else if (brswr.Equals("Safari"))
                {
                    xlWorkSheet.Cells[rowNumber, columnNumber] = "Chrome";
                }
                else
                {
                    Logger.log("check the Browser Name");
                    Assert.False(true);
                }

            }
            xlWorkBook.Save();
            closeExcel();

        }

       
    }
}
    

