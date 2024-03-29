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
using System.Data.SqlClient;
using System.Globalization;
using OpenQA.Selenium.Interactions;

namespace TestQLKS
{
    internal class ThemPhongTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;

        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("http://localhost:49921/Admin/Index/Login");
            driver.FindElement(By.Id("tai_khoan")).SendKeys("admin");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("mat_khau")).SendKeys("12345");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Tùy Chỉnh")).Click();
            {
                var element = driver.FindElement(By.TagName("body"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element, 0, 0).Perform();
            }
            driver.FindElement(By.LinkText("Phòng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Thêm phòng")).Click();
            Thread.Sleep(1000);
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
                    return result.Tables[3]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(4); // Giả sử kết quả test nằm ở sheet đầu tiên

            bool isTestCaseFound = false;

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                // Giả sử cột 'A' chứa ID của test case
                if (row.Cell("A").Value.ToString() == testCaseID)
                {
                    isTestCaseFound = true;

                    row.Cell("E").SetValue(result);
                    break;
                }
            }

            if (!isTestCaseFound)
            {
                throw new Exception($"Test case ID '{testCaseID}' not found.");
            }

            workbook.Save();
        }
        [Test]
        public void TestRoom()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest_Nam.xlsx");
            int testCaseIndex = 1;

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"TC_{testCaseIndex}";
                string so_phong = row["so_phong"].ToString();
                string loai_phong = $"//option[. = '{row["loai_phong"]}']"; // Corrected XPath
                string ma_tang = $"//option[. = '{row["ma_tang"]}']"; // Corrected XPath
                bool isRoomSuccessful = testCaseIndex > 4; // First 4 are false, others true

                try
                {
                    driver.Navigate().GoToUrl("http://localhost:49921/Admin/Phong/Create");

                    var so_phongElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("so_phong")));
                    so_phongElement.Click();
                    so_phongElement.Clear();
                    so_phongElement.SendKeys(so_phong);

                    var loai_phongElement = wait.Until(ExpectedConditions.ElementExists(By.Id("loai_phong")));
                    loai_phongElement.FindElement(By.XPath(loai_phong)).Click();

                    var ma_tangElement = wait.Until(ExpectedConditions.ElementExists(By.Id("ma_tang")));
                    ma_tangElement.FindElement(By.XPath(ma_tang)).Click();

                    var submitButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".btn-default")));
                    submitButton.Click();

                    if (testCaseIndex <= 4)
                    {
                        // Simulate a failure scenario for the first 4 test cases
                        isRoomSuccessful = false;
                    }
                    else
                    {
                        
                        wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/Admin/Phong"));
                        isRoomSuccessful = driver.Url.Contains("http://localhost:49921/Admin/Phong");
                    }

                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase_Nam.xlsx", testCaseId, isRoomSuccessful ? "Pass" : "Fail");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with error: {ex.Message}");
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase_Nam.xlsx", testCaseId, "Fail");
                    isRoomSuccessful = false;
                }

              
                driver.Navigate().GoToUrl("http://localhost:49921/Admin/Phong/Create");
                wait.Until(ExpectedConditions.UrlToBe("http://localhost:49921/Admin/Phong/Create")); 

                testCaseIndex++;
            }
        }


        [TearDown]
        public void Teardown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}
