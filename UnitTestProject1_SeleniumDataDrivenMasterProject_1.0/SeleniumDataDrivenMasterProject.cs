using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Threading;

namespace UnitTestProject1_SeleniumDataDrivenMasterProject_1._0
{
    [TestClass]
    public class SeleniumDataDrivenMasterProject
    {
        [TestMethod]
        [DataRow("lakshminarayana31.p@gmail.com", "P@ssw0rd")]
        [DataRow("lakshminarayana73.p@gmail.com","P@!sswRrd")]
        public void DataDrivenTestingUsingDataRow(string email, string pwed)
        {
            IWebDriver driver = new FirefoxDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://zixmessagecenter.com/s/welcome.jsp?b=zmc");
            driver.FindElement(By.Id("loginname")).SendKeys(email);
            driver.FindElement(By.Id("password")).SendKeys(pwed);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(10000);
            driver.Quit();
        }
        [TestMethod]
        //[DataRow("lakshminarayana31.p@gmail.com", "P@ssw0rd")]
        //[DataRow("lakshminarayana73.p@gmail.com", "P@!sswRrd")]
        [DynamicData(nameof(GetData),DynamicDataSourceType.Method)]
        public void DataDrivenTestingUsingDynamicData(string email, string pwed)
        {
            IWebDriver driver = new FirefoxDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://zixmessagecenter.com/s/welcome.jsp?b=zmc");
            driver.FindElement(By.Id("loginname")).SendKeys(email);
            driver.FindElement(By.Id("password")).SendKeys(pwed);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(10000);
            driver.Quit();
        }
        public static IEnumerable<object[]> GetData()
        {
           yield return new object[] { "lakshminarayana31.p@gmail.com", "P@ssw0rd" };
            yield return new object[] { "lakshminarayana73.p@gmail.com", "P@!sswRrd" };
        }
    }
}
