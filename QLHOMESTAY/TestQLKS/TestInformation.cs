using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;
using System.Threading;
using ExcelDataReader;
using ClosedXML.Excel;
using System.IO;
using SeleniumExtras.WaitHelpers;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;

namespace TestQLKS
{
    internal class TestInforPageTest
    {

        private IWebDriver driver;
        private WebDriverWait wait;

        public IDictionary<string, object> vars { get; private set; }
        private IJavaScriptExecutor js;
        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<string, object>();
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.CssSelector("#page > nav > div > div > div.col-xs-8.text-right.menu-1 > ul > li:nth-child(8) > a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Đăng Nhập")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("ma_kh")).SendKeys("ducnam");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("mat_khau")).SendKeys("123456");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[7]/a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Sửa thông tin")).Click();
            Thread.Sleep(1000);
        }

        [Test]
        public void TestUpdatePersonalInfo()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Downloads\\DataTest_Nam.xlsx");
            int testCaseIndex = 1; 

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"Account_{testCaseIndex}";
                string password = row["mat_khau"].ToString();
                string fullName = row["ho_ten"].ToString();
                string idNumber = row["cmt"].ToString();
                string phoneNumber = row["sdt"].ToString();
                string email = row["mail"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                string errorXPath = row["ErrorXPath"].ToString();

                try
                {
                    driver.FindElement(By.Id("mat_khau")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("mat_khau")).SendKeys(password);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("ho_ten")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("ho_ten")).SendKeys(fullName);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("cmt")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("cmt")).SendKeys(idNumber);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("sdt")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("sdt")).SendKeys(phoneNumber);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("mail")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Id("mail")).SendKeys(email);
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector(".btn")).Click();
                    Thread.Sleep(3000);
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    if (!string.IsNullOrEmpty(expectedErrorMessage))
                    {
                        var errorElement = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(errorXPath)));
                        string actualErrorMessage = errorElement.Text;
                        Assert.That(actualErrorMessage, Is.EqualTo(expectedErrorMessage), $"Test case {testCaseId} failed. Expected error message: {expectedErrorMessage}, but got: {actualErrorMessage}");
                        UpdateTestResult("C:\\Users\\dowif\\Downloads\\Testcase_Nam.xlsx", testCaseId, actualErrorMessage == expectedErrorMessage ? "Pass" : "Failed");
                    }
                    else
                    {
                        bool isNavigationSuccessful = driver.Url.Contains("http://localhost:49921/");
                        UpdateTestResult("C:\\Users\\dowif\\Downloads\\Testcase_Nam.xlsx", testCaseId, isNavigationSuccessful ? "Pass" : "Fail");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with exception: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                    UpdateTestResult("C:\\Users\\dowif\\Downloads\\Testcase_Nam.xlsx", testCaseId, "Fail");                   
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
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    return result.Tables[0]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1); // Assuming that the first worksheet is where you want to write the results

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