using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using OpenQA.Selenium.Support.UI;
using SeleniumFrameWork.Helpers;
using SeleniumFrameWork.Config;
using static SeleniumFrameWork.Base.DriverSetUp;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using AventStack.ExtentReports;
using System.Diagnostics;
using System.Drawing.Imaging;
using NUnit.Framework.Interfaces;
using System.IO;
using Ionic.Zip;

namespace SeleniumFrameWork.Base
{
    public  class HookBase:ExcelHelpers
    {
        DriverSetUp setUp = DriverSetUp.getInstance();
        ReportSetUp reportSetUp = ReportSetUp.getInstance();


        public static IWebDriver driver;
        public static ExtentReports reports;
        public static ExtentTest test;
        public static string reportPath;
        public static int itr;
        public static IWebDriver incogdriver;

        //Global wait
        public  int waitLow = 5;
        public  int waitMedium = 15;
        public  int waitHigh = 30;
        public int sleepLow = 2500;
        public int sleepMedium = 5000;
        public int sleepHigh = 12000;

        public static void NewIncogintoBrowserSetup()
        {
            string browserType = ExcelHelpers.getParameter("IncogBrwser");
            ChromeOptions options = new ChromeOptions();
            switch (browserType)
            {
                case "Chrome":

                    options.AddArguments("--disable-notifications");
                    options.AddArguments("--incognito");
                    incogdriver = new ChromeDriver(options);
                    break;

                case "Headless-Browser":
                    options.AddArguments("--disable-notifications");
                    options.AddArguments("--incognito");
                    options.AddArguments("--headless");
                    incogdriver = new ChromeDriver(options);
                    break;

                default:
                    Logger.log("Failed to Navigate through the Incognito Driver");
                    break;
            }

        }
        public static int getIteration()
        {
            /* openExcelRunConfig("RunConfigurationManager", "TestScripts");
            string testcaseID = DriverContext.getTestCaseName();
             int rowNumber = xlWorkSheet.Columns.Find(testcaseID).Cells.Row;
            int columnNumber = xlWorkSheet.Columns.Find("Iteration").Cells.Column;*/
            string iteration = (string)dataHash["Iteration"]; //casting need to be only string ( Whole method we used only string)

            int itr = Int32.Parse(iteration); // Int32 1.0 can also be passed
            return itr;
        }

        public static string Capture(string screenShotName)
        {
            var localpath = "";
            string reportPath = ReportingHelpers.reportpath();
            string timeNow = DateTime.Now.ToLongTimeString().ToString().Replace(':', '_');
            //initialize the reportsetup
            DateTime localDate = DateTime.Now;
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            Screenshot screenshot = ts.GetScreenshot();
            string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string fileName = TestContext.CurrentContext.Test.MethodName + screenShotName + DateTime.Now.Ticks + ".png";
            var sub_directorypath = reportPath + "\\ErrorScreenshots\\";
            localpath = sub_directorypath + fileName;
            if (!Directory.Exists(sub_directorypath))
            {
                Directory.CreateDirectory(sub_directorypath);
            }
            Logger.log("localpath::" + localpath);
            screenshot.SaveAsFile(localpath, ScreenshotImageFormat.Png);
            return localpath;
        }

        public static string CaptureIncogBrowser(string screenShotName)
        {
            var localpath = "";
            string reportPath = ReportingHelpers.reportpath();
            string timeNow = DateTime.Now.ToLongTimeString().ToString().Replace(':', '_');
            //initialize the reportsetup
            DateTime localDate = DateTime.Now;
            ITakesScreenshot ts = (ITakesScreenshot)incogdriver;
            Screenshot screenshot = ts.GetScreenshot();
            string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string fileName = TestContext.CurrentContext.Test.MethodName + screenShotName + DateTime.Now.Ticks + ".png";
            var sub_directorypath = reportPath + "\\ErrorScreenshots\\";
            localpath = sub_directorypath + fileName;
            if (!Directory.Exists(sub_directorypath))
            {
                Directory.CreateDirectory(sub_directorypath);
            }
            Logger.log("localpath::" + localpath);
            screenshot.SaveAsFile(localpath, ScreenshotImageFormat.Png);
            return localpath;
        }

        public static void Zip()
        {
            using (ZipFile zip = new ZipFile())
            {
                var reportSourcePath = ReportingHelpers.reportpath();
                var reportzipPath = ReportingHelpers.reportZipPath();


                zip.AddDirectory(reportSourcePath);
                zip.Save(reportzipPath + ".zip");
            }


        }
        public static void createTestCase(string TestCaseName)
        {
            test = reports.CreateTest(TestCaseName);
        }

        public static void reportLog(Status stat, string Description)
        {
            StackTrace stackTrace = new StackTrace();
            // get calling method name
            String testMethodName = stackTrace.GetFrame(1).GetMethod().Name;
            test.Log(stat, testMethodName + " :: " + Description);
            String methodName = " [ " + testMethodName + " ] ";
            

            Logger.log(methodName+Description);
            //reports.Flush();



        }
        public void switchToOriginalTab()
        {
            var originalTab = driver.SwitchTo().Window(driver.WindowHandles.First());

        }

        
        public void logTestStatus()
        {
            var status = TestContext.CurrentContext.Result.Outcome.Status;
            var stacktrace = string.IsNullOrEmpty(TestContext.CurrentContext.Result.StackTrace)
                                 ? ""
                                 : $"{TestContext.CurrentContext.Result.StackTrace}";

            var message = string.IsNullOrEmpty(TestContext.CurrentContext.Result.Message)
                              ? ""
                              : $"{TestContext.CurrentContext.Result.Message}";
            Status logStatus;

            switch (status)
            {
                case TestStatus.Failed:
                    logStatus = Status.Fail;
                    break;
                case TestStatus.Inconclusive:
                    logStatus = Status.Warning;
                    break;
                case TestStatus.Skipped:
                    logStatus = Status.Skip;
                    break;
                default:
                    logStatus = Status.Pass;
                    break;
            }

            if (status == NUnit.Framework.Interfaces.TestStatus.Failed)
            {
                reportLog(logStatus, "<b>" + "ERROR MESSAGE: " + "</b>" + message);
                reportLog(logStatus, "<b>" + "STACKTRACE: " + "</b>" + stacktrace);

                // Take a screenshot and attach it to report on failure, ideas here, can't get any of them working yet:
                string screenShotPath = HookBase.Capture("_ErrorScreenshot_");
                try
                {
                    reportLog(logStatus, "Screenshot: " + test.AddScreenCaptureFromPath(screenShotPath));

                    string screenShotpath = HookBase.CaptureIncogBrowser("_ErrorScreenshot_");
                    try
                    {
                        reportLog(logStatus, "Screenshot: " + test.AddScreenCaptureFromPath(screenShotpath));
                    }
                    catch
                    {
                        Logger.log("No IncogBrowser Errors were found");
                    }
                }
                catch(Exception Ex)
                {
                    Logger.log("Error::" + Ex);
                }
               

            }

            else
            {
                reportLog(logStatus, "Test ended with " + "<b>" + logStatus + "</b>");
            }



        }
        
        //Deleting all cookies
        public void clearCache()
        {
            driver.Manage().Cookies.DeleteAllCookies();
        }
        //handling Alerts
        public void handlingAlerts()
        {
            IAlert simpleAlert = driver.SwitchTo().Alert();
            simpleAlert.Accept();



        }
        
        public void iFramesHandling(By element)
        {

            IList<IWebElement> frameList = driver.FindElements(By.TagName("iFrame"));
            if (frameList.Count > 1)
            {
                for (int i = 0; i < frameList.Count; i++)
                {
                    driver.SwitchTo().Frame(frameList[i]);//
                    Thread.Sleep(sleepLow);
                    if (ObjectExist(element))
                    {
                        break;
                    }
                }

            }

        }

        public int generateUniqueIntegerValue(int size)
        {
            Random rm = new Random();
            string random = "";

            for(int i =0; i<size; i++)
            {
                int ran = rm.Next(9);
                random = ran + random;
            }
            int r = Int32.Parse(random);

            return r;

        }

        //Switching back to parent window this method  can use universally right wherever we need to switch to parent window 

        public void switchToParentWindow()
        {
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(sleepMedium);
        }

        //To Check Element/Object is present or existed.
        public Boolean ObjectExist(By element)
        {
            bool flag = true;
            if (driver.FindElements(element).Count == 0)
            {
                flag = false;
            }

            return flag;
        }
        //Scrolling using JavaScript Executor 

        public void JexecutorScroll(int xCoordinate, int yCoordinate)
        {
            IJavaScriptExecutor sDown = (IJavaScriptExecutor)driver;
            string windowScroll = "(window.scrollBy(" + xCoordinate + "," + yCoordinate + "))";
            sDown.ExecuteScript(windowScroll);
            Thread.Sleep(sleepLow);

        }
        //Drag and Drop
        public void dragAndDrop(IWebDriver driver, IWebElement srcElement, IWebElement dstElement)
        {

            Actions action = new Actions(driver);
            IAction DragAndDrop = action.ClickAndHold(srcElement).MoveToElement(dstElement).Release(dstElement).Build();
            DragAndDrop.Perform();
            Thread.Sleep(sleepMedium);


        }

  


        //switching windows/Tabs
        public void switchTabandWindows(string titlePage)
        {

            IList<string> list = driver.WindowHandles;

            for (int i = 0; i < list.Count(); i++)
            {
                driver.SwitchTo().Window(list[i]);
                string Title = driver.Title;
                if (Title.Contains(titlePage))
                {
                    break;

                }

            }

        }
        public void switchTabandWindowsByIndex(int titlePage)
        {

            IList<string> list = driver.WindowHandles;

            
            
                driver.SwitchTo().Window(list[titlePage]);
               
                

       
     
        }

        // Wait 
        //Time Out Wait 
        public void explicitWait(By element, int waitMin)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(waitMin));
                IWebElement ele = wait.Until(ExpectedConditions.ElementToBeClickable(element));

            }
            catch (Exception ex)
            {
                Logger.log(ex.Message);
            }
        }
            public void presenceofElement(By element, int waitMin)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(waitMin));
                    IWebElement myDynamicElement = wait.Until<IWebElement>(d => d.FindElement(element));
                }
                catch (Exception ex)
                {
                    Logger.log(ex.Message);
                }
            }

            public  void AssertElementPresent( IWebElement element)
            {
                if (!IsElementPresent(element))
                    throw new Exception(string.Format("Element is not Present Exception"));

            }

            public bool IsElementPresent(IWebElement element)
            {
                try
                {
                    bool ele = element.Displayed;
                    return true;

                }
                catch (Exception)
                {
                    return false;

                }



            }


        //using( temporary instance) closes what all gets opened. 

        public static void highLightElement(By element, IWebDriver driver)
        {
            
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].style.background='yellow'", driver.FindElement(element));

        }

        public static void scroll(By element, IWebDriver driver)
        {
            IWebElement element1 = driver.FindElement(element);
            Actions actions = new Actions(driver);
            actions.MoveToElement(element1);
            actions.Perform();
        }

        public static void scrollup(IWebDriver driver)
        {
            try
            {
                ((IJavaScriptExecutor)

                        driver).ExecuteScript("window.scrollTo(document.body.scrollHeight,0)");
            }
            catch (Exception e)
            {
                throw new Exception();
            }
        }


        public IWebElement elementToBeClickable(IWebDriver driver, By by)
        {
            try
            {
                return new WebDriverWait(driver, TimeSpan.FromSeconds(180))
                               .Until(ExpectedConditions.ElementToBeClickable(by));
            }
            catch (Exception e)
            {
                throw new Exception(e.ToString());

            }
        }
        public static void JClear(IWebDriver driver, By by)
        {
            IWebElement element = driver.FindElement(by);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].value = '';", element);
        }


        public static string getAttributeValues(IWebDriver driver, By by, string attributename)
        {
            IWebElement element = driver.FindElement(by);
            string txtdata = element.GetAttribute(attributename);
            return txtdata;
        }
        public static void jClick(IWebDriver driver, By by)
        {
            IWebElement element = driver.FindElement(by);
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].click()", element);
        }

        public static void presenceOfElement(IWebDriver driver, By by)
        {
            try
            {
                new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementIsVisible(by));
            }
            catch (Exception e)
            {
                throw new Exception(e.ToString());
            }
        }


        public static void WebElementjClick(IWebDriver driver, IWebElement element)
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].click()", element);
        }

        public void selectPicklistValue(string picklistValue)

        {
            IWebElement eventLink = driver.FindElement(By.XPath("//*[@title='" + picklistValue + "']"));
            Thread.Sleep(5000);
            WebElementjClick(driver, eventLink);
        }

        public static void JSendkeys(IWebDriver driver, By by, String text)
        {
            IWebElement element = driver.FindElement(by);
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].value= " + text + ";", element);
        }



        public static void Jscroll(By by, IWebDriver driver)
        {
            IWebElement element = driver.FindElement(by);
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }

        public static void JscrollBottom()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(0,document.body.scrollHeight)");
        }

    }

}