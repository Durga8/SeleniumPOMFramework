//comments
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SeleniumFrameWork.Base
{
    public class BasePage
    {
        public BasePage CurrentPage { get; set; }

        protected TPage GetInstance<TPage>() where TPage : BasePage, new()
        {

            {
                return (TPage)Activator.CreateInstance(typeof(TPage));
            }


        }

        public TPage As<TPage>() where TPage : BasePage
        {
            return (TPage)this;
        }
    }
}
