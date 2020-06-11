using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumFrameWork.ConfigElement
{
    [ConfigurationCollection(typeof(FrameworkElement), AddItemName = "testSetting", CollectionType = ConfigurationElementCollectionType.BasicMap)]
    public class FrameworkElementCollection:ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new FrameworkElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as FrameworkElement).Name;
        }

        public new FrameworkElement this[string type]
        {
            get
            {
                return (FrameworkElement)base.BaseGet(type);
            }
        }
    }
}
