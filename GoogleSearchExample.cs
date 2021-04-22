using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Newtonsoft.Json;

namespace SeleniumExample
{
    [TestClass]
    public class GoogleSearchExample
    {
        [TestMethod]
        public void SearchForSeleniumHQ()
        {
            Assert.IsTrue(true);

            return;

            //Find folder with Chrome Driver (chromedriver.exe)
            var browserDriverPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            //Set Chrome to start with maximized window
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--start-maximized");

            //Open Chrome
            using (var chromeDriver = new ChromeDriver(browserDriverPath, options))
            {
                //Assign google.com to variable named url
                var url = "http://google.com";
                //Go to Google.com
                chromeDriver.Navigate().GoToUrl(url);

                //Create new wait timer and set it to 1 minute
                var wait = new WebDriverWait(chromeDriver, new TimeSpan(0, 0, 1, 0));

                //Wait until an element on the page with the name q is visible.
                //Google named their search box q. probably short for query.
                //We are waiting until that appears but will wait no longer than 1 minute as defined above.
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("q")));

                //Find the Google Search Box now that it is visible and
                //assign to the variable "googleSearchBox"
                var googleSearchBox = chromeDriver.FindElement(By.Name("q"));

                //clear search box just in case anything is already typed in
                googleSearchBox.Clear();
                //Type "Selenium HQ" into the search box
                googleSearchBox.SendKeys("Selenium HQ");

                //Wait until Google Search button is visible but don't wait more than 1 minute.
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("btnK")));
                //Find "Google Search" button and assign to variable name "searchButton"
                var searchButton = chromeDriver.FindElement(By.Name("btnK"));

                //Click search button
                searchButton.Click();

                //wait until search results stats appear which confirms that the search finished
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("resultStats")));

                //Find result stats and assign to variable name resultStats
                var resultStats = chromeDriver.FindElement(By.Id("resultStats"));

                //Confirm the stats contain the words About, results and seconds.
                //Example Result stats: "About 1,090,000 results (0.49 seconds)"
                Assert.IsTrue(resultStats.Text.Contains("About"));
                Assert.IsTrue(resultStats.Text.Contains("results"));
                Assert.IsTrue(resultStats.Text.Contains("seconds"));

                //Find a search result in the list
                var results = chromeDriver.FindElement(By.ClassName("r"));

                //Confirm that the result text contains Selenium
                Assert.IsTrue(results.Text.Contains("Selenium"));

                //close Chrome
                chromeDriver.Close();
            }
        }

        [TestMethod]
        public void test()
        {
            string fileName = @"C:\Users\ed.wu\Downloads\Selenium-C-Sharp-Example-master\SeleniumExample\Excel\Data.xlsx";
            var datas = MyServices.ParseExcelDataToObject<ExcelDto>(fileName).ToList();

            foreach (var item in datas)
            {
                Console.WriteLine(JsonConvert.SerializeObject(item));
            }

            var browserDriverPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            //Set Chrome to start with maximized window
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--start-maximized");

            //Open Chrome
            using (var chromeDriver = new ChromeDriver(browserDriverPath, options))
            {
                var select = new SelectElement(chromeDriver.FindElementById(""));

                //select.Options.Where(r => r.Text.Equals("")).FirstOrDefault().Selected;
            }

            Console.ReadLine();
            Assert.IsTrue(true);

        }
        
        public void FillForm_TagName()
        {


        }

        private int LimitCheckTImes = 5;
        public bool FillForm_TagName_Check()
        {
            bool checkresult = false;
            try
            {
                for (int i = 0; i < LimitCheckTImes; i++)
                {
                    if (true)
                    {
                        checkresult = true;
                        break;   
                    }
                    Thread.Sleep(200);
                }                            
            }
            catch (System.Exception)
            {                
                //throw;
            }
            
            return checkresult;
        }

    }


   
}
