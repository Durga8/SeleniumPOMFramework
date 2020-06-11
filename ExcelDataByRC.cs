using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using xl = Microsoft.Office.Interop.Excel;

namespace SeleniumFrameWork.Helpers
{
    public class ExcelDataByRC
    {
        public static xl.Application xlAppl = null;
        public static xl.Workbooks xlworkbooks = null;
        public static xl.Workbook workbook = null;

        public static Hashtable sheets;



        //Open Excel
        public static void openExcel(string WorkbookName, string SheetName)
        {

            try
            {
                xlAppl = new xl.Application();

                //open workbook
                string sltnPath = DriverContext.getSolutionPath();
                string testDataPath = sltnPath + @"TestData\" + WorkbookName + ".xlsx";
                Thread.Sleep(2000);
                xlworkbooks = xlAppl.Workbooks;
                workbook = xlworkbooks.Open(testDataPath);
                //Access worksheet



                //Storing the Worsksheet names in Hashtable

                int count = 1;
                sheets = new Hashtable();
                foreach (xl.Worksheet sh in workbook.Sheets)
                {
                    sheets[count] = sh.Name;
                    count++;
                }

            }
            catch (Exception e)
            {
                Logger.log("error::" + e.Message);
            }
        }
        //Close Excel
        public static void closeExcel()
        {
            try
            {
                workbook.Save();
                //workbook close--->Close the connection to Workbook
                workbook.Close();
                Marshal.FinalReleaseComObject(workbook);
                workbook = null;
                //WorkBooks Close
                xlworkbooks.Close();
                Marshal.FinalReleaseComObject(xlworkbooks);
                xlworkbooks = null;

                xlAppl.Quit();
                Marshal.FinalReleaseComObject(xlAppl);
                xlAppl = null;

            }
            catch (Exception ex)
            {
                Logger.log("Error::" + ex.Message);

            }
        }



        //Read Data By Row Number and Column Name ( Get Cell Data using Column Name and by Row Number)

        public static string getCellData(string WorkbookName, string SheetName, string ColName, int rowNum)
        {

            openExcel(WorkbookName, SheetName);

            string value = string.Empty;
            int sheetValue = 0;
            int colNumber = 0;

            if (sheets.ContainsValue(SheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(SheetName))
                    {
                        sheetValue = (int)sheet.Key;

                    }
                    xl.Worksheet worksheet = null;
                    worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                    xl.Range range = worksheet.UsedRange;
                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);

                        if (colNameValue.ToLower() == ColName.ToLower())
                        {
                            colNumber = i;
                            break;

                        }
                    }

                    value = Convert.ToString((range.Cells[rowNum, colNumber] as xl.Range).Value2);
                    Marshal.FinalReleaseComObject(worksheet);
                    worksheet = null;


                }




            }
            closeExcel();
            return value;
        }

        // Write Data by Row Num and Column Name( Set Cell Data using Column Name and by Row Number)
        public static bool setCellData(string WorkbookName, string SheetName, string colName, int rowNumber, string dataValue)
        {
            openExcel(WorkbookName, SheetName);

            int sheetVal = 0;
            int colNum = 0;
            try
            {
                if (sheets.ContainsValue(SheetName))
                {
                    foreach (DictionaryEntry sheet in sheets)
                    {
                        if (sheet.Value.Equals(SheetName))
                        {
                            sheetVal = (int)sheet.Key;
                        }
                        xl.Worksheet workSheet = null;
                        workSheet = workbook.Worksheets[SheetName] as xl.Worksheet;
                        xl.Range range_01 = workSheet.UsedRange;

                        for (int i = 1; i <= range_01.Columns.Count; i++)
                        {
                            string colNameVal = Convert.ToString((range_01.Cells[1, i] as xl.Range).Value2);
                            if (colNameVal.ToLower() == colName.ToLower())
                            {
                                colNum = i;
                                break;
                            }
                        }
                        range_01.Cells[rowNumber, colNum] = dataValue;
                        workbook.Save();
                        Marshal.FinalReleaseComObject(workSheet);
                        workSheet = null;
                        closeExcel();
                    }
                }

            }
            catch
            {
                //Logger.log("Error While Writing Data in Excel::" + ex.Message);
                //return false;
            }

            return true;

        }

    }
}
