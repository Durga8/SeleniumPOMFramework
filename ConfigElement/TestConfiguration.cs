using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.ConfigElement
{
    public class TestConfiguration:ConfigurationSection
    {
        private static TestConfiguration _testConfig = (TestConfiguration)ConfigurationManager.GetSection("TestConfiguration");

        public static TestConfiguration AppSettings { get { return _testConfig; } }

        [ConfigurationProperty("testSettings")]
        public FrameworkElementCollection TestSettings { get { return (FrameworkElementCollection)base["testSettings"]; } }

    }
}
