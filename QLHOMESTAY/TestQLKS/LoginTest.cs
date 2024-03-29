using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using NUnit.Framework;
using ExcelDataReader;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using SeleniumExtras.WaitHelpers;


namespace TestQLKS
{
    [TestFixture]
    public class LoginTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;

   
        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/div/ul/li[1]/a")).Click();
        }
        private DataTable ReadTestData(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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
                    return result.Tables[0];
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(5); // Sửa lại số thứ tự worksheet nếu cần
            bool isTestCaseFound = false;

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                if (row.Cell(1).Value.ToString() == testCaseID) // Giả sử ID test case nằm ở cột 1
                {
                    isTestCaseFound = true;
                    row.Cell(6).SetValue(result); // Giả sử kết quả test được lưu ở cột 6
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
        public void Login()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Downloads\\dataTest_Tho.xlsx");
            int testCaseIndex = 1;

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"Login_{testCaseIndex}";
                string maKh = row["ma_kh"].ToString(); // Giả sử tên cột trong Excel là "ma_kh"
                string matKhau = row["mat_khau"].ToString(); // Giả sử tên cột trong Excel là "mat_khau"
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                string errorXPath = row["ErrorXPath"].ToString();

                try
                {
                    driver.FindElement(By.Id("ma_kh")).Click();
                    driver.FindElement(By.Id("ma_kh")).Clear();
                    driver.FindElement(By.Id("ma_kh")).SendKeys(maKh);
                    Thread.Sleep(100);

                    driver.FindElement(By.Id("mat_khau")).Click();
                    driver.FindElement(By.Id("mat_khau")).Clear();
                    driver.FindElement(By.Id("mat_khau")).SendKeys(matKhau);
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".btn-primary")).Click();
                    Thread.Sleep(100);

                    // Kiểm tra trường hợp có thông báo lỗi xuất hiện hay không
                    if (!string.IsNullOrEmpty(expectedErrorMessage))
                    {
                        var errorElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(errorXPath)));
                        string actualErrorMessage = errorElement.Text;
                        Assert.That(actualErrorMessage, Is.EqualTo(expectedErrorMessage), $"Test case {testCaseId} failed. Expected error message: {expectedErrorMessage}, but got: {actualErrorMessage}");

                        // Cập nhật kết quả thành công hoặc thất bại vào file test cases
                        UpdateTestResult("C:\\Users\\dowif\\Downloads\\testCase_Tho.xlsx", testCaseId, actualErrorMessage == expectedErrorMessage ? "Pass" : "Failed");
                    }
                    else
                    {
                        // Trường hợp không có lỗi và chuyển trang dự kiến
                        wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/")); // Chờ cho đến khi URL trang chủ xuất hiện
                        Assert.That(driver.Url, Does.Contain("http://localhost:49921/"), "Không quay lại trang chủ");

                        // Cập nhật kết quả thành công vào file test cases
                        UpdateTestResult("C:\\Users\\dowif\\Downloads\\testCase_Tho.xlsx", testCaseId, "Pass");
                    }
                    // Reset trạng thái cho lần test tiếp theo nếu cần
                }
                catch (Exception ex)
                {
                    // Nếu có lỗi xảy ra, cập nhật kết quả thất bại vào file test cases
                    UpdateTestResult("C:\\Users\\dowif\\Downloads\\testCase_Tho.xlsx", testCaseId, "Fail");
                    // Ghi lại thông tin lỗi nếu cần
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with error: {ex.Message}");
                }
                testCaseIndex++;
            }
        }

        [TearDown]
        protected void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}