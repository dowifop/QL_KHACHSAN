using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Data;
using ExcelDataReader;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
namespace TestQLKS
{
    public class RegisterTests
    {
        private IWebDriver driver;
        private WebDriverWait wait;

        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
          


            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.CssSelector("#page > nav > div > div > div.col-xs-8.text-right.menu-1 > ul > li:nth-child(8) > a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("//*[@id=\"page\"]/nav/div/div/div[2]/ul/li[8]/div/ul/li[2]/a")).Click();
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
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(5); // Giả sử kết quả test nằm ở sheet đầu tiên

            bool isTestCaseFound = false;

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                // Giả sử cột 'A' chứa ID của test case
                if (row.Cell("A").Value.ToString() == testCaseID)
                {
                    isTestCaseFound = true;
                    // Cập nhật cột 'G' với kết quả, đảm bảo rằng đây là cột đúng trong file của bạn
                    row.Cell("F").SetValue(result);
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
        public void RegisterWithTestData()
        {
            // Đọc dữ liệu test từ file Excel
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            int testCaseIndex = 1;
            foreach (DataRow row in testData.Rows)
            {
                // Define testCaseId at the beginning of the loop
                string testCaseId = $"{testCaseIndex}";

                // Lấy thông tin từ datatest
                string ma_kh = row["ma_kh"].ToString();
                string mat_khau = row["mat_khau"].ToString();
                string ho_ten = row["ho_ten"].ToString();
                string cmt = row["cmt"].ToString();
                string sdt = row["sdt"].ToString();
                string mail = row["mail"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                string errorXPath = row["ErrorXPath"].ToString();
                try
                {
                    // Điền thông tin vào form đăng ký
                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ma_kh")));
                    driver.FindElement(By.Id("ma_kh")).Clear();
                    driver.FindElement(By.Id("ma_kh")).SendKeys(ma_kh);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("mat_khau")).Clear();
                    driver.FindElement(By.Id("mat_khau")).SendKeys(mat_khau);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("ho_ten")).Clear();
                    driver.FindElement(By.Id("ho_ten")).SendKeys(ho_ten);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("cmt")).Clear();
                    driver.FindElement(By.Id("cmt")).SendKeys(cmt);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("sdt")).Clear();
                    driver.FindElement(By.Id("sdt")).SendKeys(sdt);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("mail")).Clear();
                    driver.FindElement(By.Id("mail")).SendKeys(mail);
                    Thread.Sleep(1000);
                    // Submit form đăng ký
                    driver.FindElement(By.CssSelector("input[type='submit']")).Click();

                    // Kiểm tra trường hợp có thông báo lỗi xuất hiện hay không
                    if (!string.IsNullOrEmpty(expectedErrorMessage))
                    {
                        var errorElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(errorXPath)));
                        string actualErrorMessage = errorElement.Text;
                        Assert.That(actualErrorMessage, Is.EqualTo(expectedErrorMessage), $"Test case {testCaseId} failed. Expected error message: {expectedErrorMessage}, but got: {actualErrorMessage}");

                        // Cập nhật kết quả thành công hoặc thất bại vào file test cases
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, actualErrorMessage == expectedErrorMessage ? "Pass" : "Failed");
                    }
                    else
                    {
                        // Trường hợp không có lỗi và chuyển trang dự kiến
                        wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/")); // Chờ cho đến khi URL trang chủ xuất hiện
                        Assert.That(driver.Url, Does.Contain("http://localhost:49921/"), "The home page was not reached after registration.");

                        // Cập nhật kết quả thành công vào file test cases
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Pass");
                    }
                }
                catch (Exception ex)
                {
                    // Nếu có lỗi xảy ra, cập nhật kết quả thất bại vào file test cases
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                    // Ghi lại thông tin lỗi nếu cần
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with error: {ex.Message}");
                }
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
