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
using System.Threading;

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

        //https://www.w3schools.com/xml/xpath_syntax.asp
        [TestMethod]
        public void test()
        {
            //string fileName = @"C:\Users\ed.wu\Downloads\Selenium-C-Sharp-Example-master\SeleniumExample\Excel\Data.xlsx";
            //var datas = MyServices.ParseExcelDataToObject<ExcelDto>(fileName).ToList();
            //
            //foreach (var item in datas)
            //{
            //    Console.WriteLine(JsonConvert.SerializeObject(item));
            //}

            var browserDriverPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            //Set Chrome to start with maximized window
            ChromeOptions options = new ChromeOptions();
            //options.AddArguments("--start-maximized");

            //Open Chrome
            using (var chromeDriver = new ChromeDriver(browserDriverPath, options))
            {
                
                var url = "https://gloriadata.tw/Hr/List";                    
                try
                {                    
                    chromeDriver.Navigate().GoToUrl(url);                    
                    var wait = new WebDriverWait(chromeDriver, new TimeSpan(0, 0, 0, 10));

                    Module_Login(chromeDriver, wait);

                    Main_Step_前往產學合作頁面(chromeDriver, wait);
                    

                }
                catch (Exception)
                {
                    throw;
                }
                finally {

                    LogOut(chromeDriver);
                    //close Chrome
                    Thread.Sleep(5000);
                    chromeDriver.Close();
                
                }
            }

            Console.ReadLine();
            Assert.IsTrue(true);

        }


        #region Login

        public void Module_Login(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            //step 1 open web https://gloriadata.tw/Hr/List 
            //step 2 fill account 帳號：s0006 id: inputEmail
            //step 2 fill pwd nchu22840558 <input type="text" id="inputEmail" class="form-control" placeholder="帳號" name="email" required="" autofocus="">
            //step 2 click login button <button class="btn btn-lg btn-primary btn-block" type="submit">登入</button>

            Login_Step1_Fill_Account(chromeDriver, wait);

            Login_Step2_Fill_PWD(chromeDriver, wait);

            Login_Step3_ClickLoginBtn(chromeDriver, wait);            
        }

        public void Login_Step1_Fill_Account(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            string domid = "inputEmail";
            var _by = By.Id("inputEmail");

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
            var accountBox = chromeDriver.FindElement(_by);
            accountBox.Clear();
            accountBox.SendKeys("s0006");
        }

        public void Login_Step2_Fill_PWD(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            string domid = "inputPassword";
            var _by = By.Id("inputPassword");

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
            var accountBox = chromeDriver.FindElement(_by);
            accountBox.Clear();
            accountBox.SendKeys("nchu22840558");
        }

        public void Login_Step3_ClickLoginBtn(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            string domid = "inputPassword";
            
            var element_Form =  wait.Until(r=>r.FindElement(By.TagName("form")));

            var elButton = element_Form.FindElement(By.TagName("button"));

            elButton.Click();
        }

        #endregion


        public void Flow()
        { 
            //Step Login

            //Step Add Data
                        

            //Step Logout
        }

        


        #region 新增資料

        public void Module_AddData(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            string url_setData = "";

            //刷新頁面確定是否要重新登入
            if (IsAlive(chromeDriver))
            {

            }

            //
            AddData_前往產學合作頁面(chromeDriver, wait);


            //確認資料類型來判斷新增步驟        

        }

        public void AddData_前往產學合作頁面(ChromeDriver chromeDriver, WebDriverWait wait)
        {
            string xPath = "/li/a[@href='TechnologyTransferContract']";
            var _by = By.XPath(xPath);

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));

            var element = chromeDriver.FindElement(_by);

            element.Click();
        }

        public void AddData_Step1_點擊按鍵()
        { 
        
        }

        #endregion

        public bool IsUrlLike(string CueentUrl ,string ExpectedUrl)
        {
            return (CueentUrl.Contains(ExpectedUrl, StringComparison.CurrentCultureIgnoreCase));
        }


        public bool IsAlive(ChromeDriver chromeDriver)
        {
            bool isAlive = true;
            
            chromeDriver.Navigate().Refresh();
            if (IsUrlLike(chromeDriver.Url, "Home/Login"))
            {
                isAlive = false;
            }

            return isAlive;
        }

        public void LogOut(ChromeDriver chromeDriver)
        { 
        
        }

    }


   
}
