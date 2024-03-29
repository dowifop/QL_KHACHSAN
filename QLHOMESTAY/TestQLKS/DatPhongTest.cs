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

namespace TestQLKS
{
    internal class DatPhongTest
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
            driver.FindElement(By.CssSelector("#page > nav > div > div > div.col-xs-8.text-right.menu-1 > ul > li:nth-child(8) > a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Đăng Nhập")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("ma_kh")).SendKeys("binh");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("mat_khau")).SendKeys("123456");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[3]/a")).Click();
            Thread.Sleep(1000);
            IWebElement dateStartElement = driver.FindElement(By.Id("datestart"));
            Thread.Sleep(1000);
            dateStartElement.Click();
            Thread.Sleep(1000);
            IWebElement dateEndElement = driver.FindElement(By.Id("dateend"));
            Thread.Sleep(1000);
            dateEndElement.Click();
            Thread.Sleep(1000);
            dateEndElement.SendKeys(Keys.Control + "a");
            dateEndElement.SendKeys(Keys.Delete);
            dateEndElement.SendKeys("30/03/2024");
            dateEndElement.SendKeys(Keys.Enter);
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn")).Click();
        }


        [Test]
        public void DatPhong()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            int testCaseIndex = 1;

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"DP_0{testCaseIndex}"; 
                string datestart = row["datestart"].ToString();
                string dateend = row["dateend"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                bool isDateSelectionSuccessful = false;
                string actualErrorMessage = "";
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
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
                    driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[2]/form/div[2]/div[3]/input")).Click();

                    // Kiểm tra xem alert có xuất hiện hay không
                    bool isAlertPresent = IsAlertPresent(wait);
                    if (isAlertPresent)
                    {
                        IAlert alert = driver.SwitchTo().Alert();
                        actualErrorMessage = alert.Text;
                        alert.Accept();
                        isDateSelectionSuccessful = actualErrorMessage.Contains(expectedErrorMessage);
                    }
                    else
                    {
                        isDateSelectionSuccessful = PerformNextActions(driver, wait);
                        // Chỉ khi thực hiện các bước thành công, mới kiểm tra dữ liệu trong DB
                        if (isDateSelectionSuccessful)
                        {
                            bool startDateSuccess = DateTime.TryParseExact(datestart, "dd/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate);
                            bool endDateSuccess = DateTime.TryParseExact(dateend, "dd/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate);
                            // Kiểm tra ngày đặt phòng trong cơ sở dữ liệu
                            VerifyBookingDatesInDatabase(testCaseId, startDate, endDate);
                        }
                    }
                    // Cập nhật kết quả test
                    string result = isDateSelectionSuccessful ? "Pass" : "Fail";
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, result);
                }
                catch (Exception ex)
                {
                    
                    Console.WriteLine($"Test thất bại cho ID test case: {testCaseId} với lỗi: {ex.Message}");
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                }

                testCaseIndex++;
            }
        }

        private bool IsAlertPresent(WebDriverWait wait)
        {
            try
            {
                wait.Until(ExpectedConditions.AlertIsPresent());
                return true;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        private bool PerformNextActions(IWebDriver driver, WebDriverWait wait)
        {
            try
            {
                driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[2]/div/table/tbody/tr[2]/td[6]/button")).Click();
                Thread.Sleep(1000); 
                driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[2]/form/div[2]/div[4]/a")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[1]/div/div[2]/form/div/div[8]/div/input")).Click();

                bool hasNavigated = wait.Until(d => d.Url.Contains("http://localhost:49921/Home/Result"));
                return hasNavigated;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Một ngoại lệ xảy ra khi thực hiện các bước tiếp theo: {ex.Message}");
                return false; 
            }
        }
        private void VerifyBookingDatesInDatabase(string testCaseId, DateTime dateStart, DateTime dateEnd)
        {
            string connectionString = "data source=.;initial catalog=dataQLKS;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework";

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                // Truy vấn kiểm tra sự tồn tại của ngày vào và ngày ra trong bất kỳ bản ghi nào
                var query = @"
                        SELECT COUNT(*)
                        FROM tblPhieuDatPhong
                        WHERE CONVERT(date, ngay_vao) = @dateStart AND CONVERT(date, ngay_ra) = @dateEnd";
                using (var command = new SqlCommand(query, connection))
                {
                    // Thêm tham số để tránh SQL Injection
                    command.Parameters.AddWithValue("@dateStart", dateStart.Date);
                    command.Parameters.AddWithValue("@dateEnd", dateEnd.Date);
                    // Thực thi truy vấn và lấy kết quả
                    int count = (int)command.ExecuteScalar();

                    // Kiểm tra xem có bản ghi nào khớp với điều kiện không
                    bool datesExist = count > 0;

                    Console.WriteLine($"Debugging: {testCaseId}");
                    Console.WriteLine($"Checking for dates in database: '{dateStart}' to '{dateEnd}', Found: {datesExist}");

                    // Cập nhật kết quả kiểm tra
                    if (datesExist)
                    {
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Pass");
                    }
                    else
                    {
                        Console.WriteLine($"No booking data found in database for dates: {dateStart} to {dateEnd}");
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                    }
                }
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
                    return result.Tables[3];
                }   
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(3);
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
