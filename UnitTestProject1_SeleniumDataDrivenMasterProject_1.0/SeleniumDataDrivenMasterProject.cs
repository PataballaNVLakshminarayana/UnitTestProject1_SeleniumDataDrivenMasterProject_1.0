using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using TestContext = Microsoft.VisualStudio.TestTools.UnitTesting.TestContext;

namespace UnitTestProject1_SeleniumDataDrivenMasterProject_1._0
{
    [TestClass]
    public class SeleniumDataDrivenMasterProject
    {
        //IWebDriver driver;
        //ExtentReports extent;
        //ExtentTest test;
        //[SetUp]
        //public void Report()
        //{
        //     extent = new ExtentReports();
        //    var htmlreporte = new ExtentHTMLReporter();
        //}
        [TestMethod]
        [DataRow("lakshminarayana31.p@gmail.com", "P@ssw0rd")]
        [DataRow("lakshminarayana73.p@gmail.com", "P@!sswRrd")]
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
        [TestMethod]
        
        [DynamicData(nameof(ReadExcel), DynamicDataSourceType.Method)]
        public void DataDrivenTestingUsingExcelSheet(string email, string pwed)
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
        public static IEnumerable<object[]> ReadExcel()
        {
            //create worksheet object
            using (ExcelPackage excelpckge = new ExcelPackage(new System.IO.FileInfo("TestDataSet.xlsx")))
            {
                ExcelWorksheet excelsheet = excelpckge.Workbook.Worksheets["Sheet1"];
                int rowcount = excelsheet.Dimension.End.Row;
                for (int i = 1; i <= rowcount; i++)
                {
                    yield return new object[]
                    {
                        excelsheet.Cells[i,1].Value?.ToString().Trim(),
                        excelsheet.Cells[i,2].Value?.ToString().Trim()
                    };
                }
            }
        }
        
        //[DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML", "|DataDirectory|\\XMLTestDataSet.xml", "logins",DataAccessMethod.Sequential)]
        //[Test, TestCaseSource(nameof(GetTestData))]
        //[TestMethod]
        //public void MyTeDataDrivenTestingUsingXMLFile(Dictionary<string, string> testData)
        //{
        //    IWebDriver driver = new FirefoxDriver();
        //    driver.Manage().Window.Maximize();
        //    driver.Navigate().GoToUrl("https://zixmessagecenter.com/s/welcome.jsp?b=zmc");
        //    driver.FindElement(By.Id("loginname")).SendKeys("Username");
        //    driver.FindElement(By.Id("password")).SendKeys("Password");
        //    driver.FindElement(By.XPath("//input[@type='submit']")).Click();
        //    Thread.Sleep(10000);
        //    driver.Quit();
        //}
        //public List<Dictionary<string, string>> ReadTestData(string filePath)
        //{
        //    var testData = new List<Dictionary<string, string>>();

        //    XmlDocument xmlDoc = new XmlDocument();
        //    xmlDoc.Load(@"D:\Practice_Final\SeleniumWithC#\UnitTestProject1_SeleniumDataDrivenMasterProject_1.0\UnitTestProject1_SeleniumDataDrivenMasterProject_1.0\XMLTestDataSet.xml"); // Load the XML file

        //    // Select all the <TestCase> nodes
        //    XmlNodeList testCaseNodes = xmlDoc.GetElementsByTagName("//logins/credentials");

        //    foreach (XmlNode testCase in testCaseNodes)
        //    {
        //         Dictionary<string, string> testCaseData = new Dictionary<string, string>();
        //        XmlNode username = testCase.SelectSingleNode("Username");
        //        XmlNode password = testCase.SelectSingleNode("Password");
        //        // Read the <Username> and <Password> from the XML file
        //        //string username = testCase["Username"].InnerText;
        //        //string password = testCase["Password"].InnerText;

        //        // Add the values to the dictionary
        //        //testCaseData.Add("Username", username);
        //        //testCaseData.Add("Password", password);

        //        // Add the dictionary to the list
        //        testData.Add(testCaseData);
        //    }

        //    return testData;
        //}
        //public static IEnumerable<Dictionary<string, string>> GetTestData()
        //{
        //    var data = new SeleniumDataDrivenMasterProject().ReadTestData("XMLTestDataSet.xml");
        //    foreach (var testCase in data)
        //    {
        //        yield return testCase;
        //    }
        //}
        //BROWSER NAVIGATIONS IN SELENIUM WITH C#
        [TestMethod]
        public void BrowserNavigationGOTO()
        {
            IWebDriver driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("https://zixmessagecenter.com/s/welcome.jsp?b=zmc");
            Thread.Sleep(10000);
            driver.Url = "https://accounts.google.com/v3/signin/identifier?ifkv=AcMMx-c18DFtZscqEm4cQBsuA37Y2zBlSp2HtsYyEjjt-i3Gx7ENJhPxUjdxrg6VHLMgc2U-Jun0Vw&service=mail&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S35443072%3A1733114961883360&ddm=1";
            Thread.Sleep(10000);
            driver.Navigate().Back();
            Thread.Sleep(10000);
            driver.Navigate().Forward();
            Thread.Sleep(10000);
            driver.Navigate().Refresh();
            Thread.Sleep(10000);
            driver.Quit();
        }
        [TestMethod]
        public void ManageBrowserWindows()
        {
            IWebDriver Fire_driver = new FirefoxDriver();
            Fire_driver.Navigate().GoToUrl("https://zixmessagecenter.com/s/welcome.jsp?b=zmc");
            Thread.Sleep(10000);
            Fire_driver.Manage().Window.FullScreen();
            Thread.Sleep(10000);
            Fire_driver.Manage().Window.Maximize();
            Thread.Sleep(10000);
            Fire_driver.Manage().Window.Minimize();
            Thread.Sleep(10000);
            Fire_driver.Manage().Window.Position = new Point(400, 200);
            Point pointer = Fire_driver.Manage().Window.Position;
            Console.WriteLine(pointer);
            Thread.Sleep(10000);
            Fire_driver.Manage().Window.Size = new Size(600, 400);
            Size size= Fire_driver.Manage().Window.Size;
            Console.WriteLine(size);
            Thread.Sleep(10000);
            Fire_driver.Quit();
        }
        [TestMethod]
        public void FileUploadToBrowser()
        {
            IWebDriver fire_driver = new FirefoxDriver();
            string userProfileFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string filePath = Path.Combine(userProfileFolder, "Desktop", "SeleniumWithC#.docx");
            fire_driver.Navigate().GoToUrl("https://www.ilovepdf.com/word_to_pdf");
            IWebElement fileuploads = fire_driver.FindElement(By.Id("pickfiles"));
            fileuploads.Click();
            //string filepath = @"C:\Users\lakshminarayana.B2BSOFTECH\Desktop\SeleniumWithC#.docx";
            fileuploads.SendKeys(filePath);
            IWebElement subminting =fire_driver.FindElement(By.XPath("//button[@id='processTask']"));
            subminting.Submit();
            Thread.Sleep(10000);
            fire_driver.Quit();

        }
    }
    [TestClass]
    public class SeleniumDataDrivenMasterProject_Parllel
    {
        IWebDriver fire_driver= new FirefoxDriver();;
        [TestMethod]
        public void MyTestMethodFailed()
        {
            ProjectofPOM("https://zixmessagecenter.com/s/welcome.jsp?b=zmc", "lakshminarayana73.p@gmail.com", "Naru@0543");
           
        }
        [TestMethod]
        public void MyTestMethodPassed()
        {
            ProjectofPOM("https://zixmessagecenter.com/s/welcome.jsp?b=zmc", "support@geniusdoc.com", "Sunil7231982@");

        }
        public void ProjectofPOM(string url, string username, string _password)
        {
            
            fire_driver.Navigate().GoToUrl(url);
            fire_driver.FindElement(By.Id("loginname")).SendKeys(username);
            fire_driver.FindElement(By.Id("password")).SendKeys(_password);
            fire_driver.FindElement(By.XPath("//input[@name='login']")).Click();
            Thread.Sleep(10000);
            fire_driver.Quit();
        }
    }

}
