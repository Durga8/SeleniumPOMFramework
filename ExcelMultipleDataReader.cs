using SeleniumFrameWork.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SeleniumFrameWork.Helpers
{
    public class ExcelMultipleDataReader

    {
        public static Microsoft.Office.Interop.Excel.Application xlApp;

        public static Microsoft.Office.Interop.Excel.Workbook xlWorkBook_01;
        public static Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet_01;
        public static Hashtable dataSet = new Hashtable();


        //Last column should be Count Key in the excelWorkBook( Condition to iterate all the Rows to fetch different data each time)
        //1) Open the Excel
        public static void openExcel(string workbookName, string SheetName)
        {
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
               
                //open workbook
                string sltnPath = DriverContext.getSolutionPath();
                string testDataPath = sltnPath + @"TestData\" + workbookName + ".xlsx";
                Thread.Sleep(2000);
                xlWorkBook_01 = xlApp.Workbooks.Open(testDataPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                

                //Access worksheet

                xlWorkSheet_01 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook_01.Worksheets[SheetName];
            }
            catch (Exception e)
            {
                Logger.log("error::" + e.Message);
            }
        }
        //1)a Store all the values in a datatable
        public static void storedataTable()
        {
            try
            {
                //dataSet = new Hashtable();
                dataSet.Clear();
                int columNumber = xlWorkSheet_01.Columns.Find("Count").Cells.Column;
                string testKeyCounter = xlWorkSheet_01.Cells[2,columNumber].Text.ToString();
                int count = Int32.Parse(testKeyCounter);
                int rowCount = xlWorkSheet_01.UsedRange.Rows.Count;
                Logger.log("RowCount::" + rowCount);
                for(int i=2; i < columNumber; i++)
                {
                    
                    string key = xlWorkSheet_01.Cells[1,i].Text.ToString();
                    
                    string value = xlWorkSheet_01.Cells[count,i].Text.ToString();
                    
                    dataSet.Add(key, value);
                   
                    
                }
                if (count >= rowCount)
                {
                    xlWorkSheet_01.Cells[2, columNumber] = 2;

                }

                else
                {
                    xlWorkSheet_01.Cells[2, columNumber] = count + 1;
                }
               
                xlWorkBook_01.Save();

            }
            catch (Exception e)
            {
                Logger.log("error::" + e.Message);
            }
            finally
            {
                closeExcel();
            }
        }
                
        //2) Pick the first column data, while executing for the first time and 2nd data second time and so on

        //3) close the excel
        public static void closeExcel()
        {
            try
            {
                xlWorkBook_01.Close();
                xlApp.Quit();
            }
            catch(Exception e)
            {
                Logger.log("Error" + e.Message);
            }

        }

        //Writing Data to the Excel Multiple Data Reader

        public static void setMulData(string workbookName, string sheetName, string columnName, string dataValue)
        {
            try
            {
                openExcel(workbookName, sheetName);
                int columnNum = xlWorkSheet_01.Columns.Find(columnName).Cells.Column;
                int rowcount = xlWorkSheet_01.UsedRange.Rows.Count;
                string iteration = ExcelMultipleDataReader.getMulParameter("Iteration");
                int iterationCount = Int32.Parse(iteration);
                if (iterationCount <= rowcount)
                {
                    xlWorkSheet_01.Cells[iterationCount, columnNum] = dataValue;

                }

                xlWorkBook_01.Save();


            }
            catch (Exception e)
            {
                Logger.log("Error While Writing the Data::" + e.Message);
            }
            finally
            {
                closeExcel();
            }
        }

        public static void getMulData(string workbookName, string SheetName)
        {
            openExcel(workbookName, SheetName);
            storedataTable();
        }
        
        public static string getMulParameter(string Key)
        {
            string parameter = (string)dataSet[Key];

            Logger.log("Parameter::" + parameter);

            return parameter;
        }
    }
}
