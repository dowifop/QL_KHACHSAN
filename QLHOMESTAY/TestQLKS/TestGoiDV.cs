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
    public class TestGoiDV
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        int initialTonKho = 0;
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
            driver.FindElement(By.LinkText("Quản lý phòng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Chọn dịch vụ")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary:nth-child(1)")).Click();
            Thread.Sleep(1000);
            initialTonKho = int.Parse(driver.FindElement(By.CssSelector("#dataTable > tbody > tr:nth-child(6) > td:nth-child(4)")).Text); // Sửa selector này để phù hợp với mã HTML của bạn
            driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr[6]/td[6]/a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("so_luong")).Click();
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
                    return result.Tables[2]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(3); // Giả sử kết quả test nằm ở sheet đầu tiên

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
        public void CallServiceWithTestData()
        {
            // Đọc dữ liệu test từ file Excel
            var testData = ReadTestData("C:\\BDCLPM\\DataTest_Nam.xlsx");
            int testCaseIndex = 1;
            foreach (DataRow row in testData.Rows)
            {
                // Define testCaseId at the beginning of the loop
                string testCaseId = $"KH_0{testCaseIndex}";

                // Lấy thông tin từ datatest
                string soluong = row["so_luong"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                string actualErrorMessage = "";
                bool isTestSuccessful = true;
                try
                {

                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("so_luong")));
                    driver.FindElement(By.Id("so_luong")).Clear();
                    driver.FindElement(By.Id("so_luong")).SendKeys(soluong);
                    Console.WriteLine($"Entered quantity: {soluong}");

                    // Submit form đăng ký
                    driver.FindElement(By.CssSelector("#popup .btn-primary")).Click();
                    Console.WriteLine("Clicked confirm button on the popup.");
                    // Kiểm tra xem có thông báo lỗi hay không
                    try
                    {
                        IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
                        actualErrorMessage = alert.Text;
                        alert.Accept();
                        Console.WriteLine($"Alert present with message: {actualErrorMessage}");
                        if (!actualErrorMessage.Contains(expectedErrorMessage))
                        {
                            isTestSuccessful = false;
                            UpdateTestResult("C:\\BDCLPM\\Testcase_Nam.xlsx", testCaseId, "Fail");
                        }
                    }
                    catch (WebDriverTimeoutException)
                    {
                        Console.WriteLine("No alert present, proceeding to verify inventory count.");
                        // Nếu không có alert, kiểm tra số lượng tồn kho sau khi đặt
                        wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.Id("popup"))); // Đảm bảo rằng popup đã đóng
                        int finalTonKho = int.Parse(driver.FindElement(By.CssSelector("#dataTable > tbody > tr:nth-child(6) > td:nth-child(4)")).Text);
                        int bookedQuantity = int.Parse(soluong);
                        isTestSuccessful = (initialTonKho - bookedQuantity == finalTonKho);
                        Console.WriteLine($"Inventory before: {initialTonKho}, Inventory after: {finalTonKho}, Booked quantity: {bookedQuantity}, Test result: {isTestSuccessful}");
                    }

                    // Cập nhật kết quả test
                    if (isTestSuccessful)
                    {
                        UpdateTestResult("C:\\BDCLPM\\Testcase_Nam.xlsx", testCaseId, "Pass");
                    }
                }
                catch (Exception ex)
                {
                    // Nếu có lỗi xảy ra, cập nhật kết quả thất bại vào file test cases
                    UpdateTestResult("C:\\BDCLPM\\Testcase_Nam.xlsx", testCaseId, "Fail");
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
