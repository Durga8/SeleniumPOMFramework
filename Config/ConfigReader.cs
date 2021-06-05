using SeleniumFrameWork.Base;
using SeleniumFrameWork.ConfigElement;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;


namespace SeleniumFrameWork.Config
{
    public class ConfigReader:HookBase
    {
        public static void SetFrameworkSettings(string settingType)
        {
            Settings.AUT = TestConfiguration.AppSettings.TestSettings[settingType].AUT;
            Settings.Name = TestConfiguration.AppSettings.TestSettings[settingType].Name;
            Settings.Browser = TestConfiguration.AppSettings.TestSettings[settingType].Browser;
            Settings.TestType = TestConfiguration.AppSettings.TestSettings[settingType].TestType;
            Settings.AppConnectionString = TestConfiguration.AppSettings.TestSettings[settingType].AppConnectionString;
        }
    }
}
