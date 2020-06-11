using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Safari;
using SeleniumFrameWork.Config;
using SeleniumFrameWork.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.Base
{
    public class DriverSetUp
    { 
        private static DriverSetUp driverSetUpInstance = null;

        private static IWebDriver driver;
       

        public DriverSetUp()
        {

        }
        public static DriverSetUp getInstance()
        {

            if (driverSetUpInstance == null)
            {

                driverSetUpInstance = new DriverSetUp();


            }
            return driverSetUpInstance;

        }



        public IWebDriver getWebDriver()
        {
            
            ChromeOptions options = new ChromeOptions();
            InternetExplorerOptions caps = new InternetExplorerOptions();
            caps.IgnoreZoomLevel = true;
            caps.EnableNativeEvents = false;
            caps.InitialBrowserUrl = "http://localhost";
            caps.UnhandledPromptBehavior = UnhandledPromptBehavior.Accept;
            caps.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            caps.EnablePersistentHover = true;

             string browserType = ExcelHelpers.getParameter("Browser");

            switch (browserType)
            {
                case "Chrome":
                    options.AddArguments("--disable-notifications");

                    if (!string.IsNullOrEmpty(browserType))
                        driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);
                    else
                        driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);

                    break;
                case "IE":
                    //set capability                 
                    driver = new InternetExplorerDriver(caps);
                    break;

                case "Safari":
                    driver = new SafariDriver();
                    break;

                case "Headless-Chrome":
                    //Headless ChromeBrowser
                    options.AddArguments("--disable-notifications");
                    options.AddArguments("--headless");
                    driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);
                    break;

                default:
                    Logger.log("No Broswer Found");
                    break;

            }

            return driver;


        }


    }
}
