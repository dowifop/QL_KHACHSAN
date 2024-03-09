using ClosedXML.Excel;
using ExcelDataReader;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestQLKS
{
    public class TestTimPhong
    {
        private IWebDriver driver;
        private WebDriverWait wait;

        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[3]/a")).Click();
        }

        [Test]
        public void timphong()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            int testCaseIndex = 1;

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"TP{testCaseIndex}";
                string datestart = row["datestart"].ToString();
                string dateend = row["dateend"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                bool isDateSelectionSuccessful = true;
                string actualErrorMessage = "";

                try
                {

                    IWebElement dateStartElement = driver.FindElement(By.Id("datestart"));
                    dateStartElement.Click();
                    dateStartElement.SendKeys(Keys.Control + "a");
                    dateStartElement.SendKeys(Keys.Delete);


                    dateStartElement.SendKeys(datestart);
                    dateStartElement.SendKeys(Keys.Enter);


                    IWebElement dateEndElement = driver.FindElement(By.Id("dateend"));
                    dateEndElement.Click();
                    dateEndElement.SendKeys(Keys.Control + "a");
                    dateEndElement.SendKeys(Keys.Delete);


                    dateEndElement.SendKeys(dateend);
                    dateEndElement.SendKeys(Keys.Enter);


                    driver.FindElement(By.CssSelector(".btn")).Click();


                    try
                    {
                        IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
                        actualErrorMessage = alert.Text;
                        alert.Accept();
                        isDateSelectionSuccessful = actualErrorMessage.Contains(expectedErrorMessage);
                    }
                    catch (NoAlertPresentException)
                    {
                        wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/"));
                        isDateSelectionSuccessful = driver.Url.Contains("http://localhost:49921/");
                    }


                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, isDateSelectionSuccessful ? "Pass" : "Fail");

                }
                catch (Exception ex)
                {

                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with exception: {ex.Message}");
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                    isDateSelectionSuccessful = false;
                }

                testCaseIndex++;
            }
        }


        private DataTable ReadTestData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                    return result.Tables[2];
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(6);

            bool isTestCaseFound = false;
            foreach (IXLRow row in worksheet.Rows())
            {
                if (row.Cell(1).Value.ToString().Equals(testCaseID))
                {
                    isTestCaseFound = true;
                    row.Cell("F").Value = result;
                    break;
                }
            }

            if (!isTestCaseFound)
            {
                throw new Exception($"Test case ID '{testCaseID}' not found in the Excel file.");
            }

            workbook.Save();
        }
        [TearDown]
        public void Teardown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}
