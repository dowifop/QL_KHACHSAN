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
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.SqlClient;

namespace TestQLKS
{
    public class TestXoaKH
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private string connectionString = "data source=.;initial catalog=dataQLKS;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework";
        private string currentUrl;
        private bool existsInDatabase;
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
            driver.FindElement(By.CssSelector(".nav-item:nth-child(6) > .nav-link")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Khách Hàng")).Click();
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
                    return result.Tables[1]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(2); // Giả sử kết quả test nằm ở sheet đầu tiên

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
        public void DeleteAndVerifyCustomer()
        {
            // Đăng nhập và đi đến trang khách hàng

            // Xóa khách hàng và kiểm tra
            DeleteCustomer();
            string testCaseId1 = "KH_01";
            string filePath = "C:\\BDCLPM\\Testcase_Nam.xlsx";
            UpdateTestResult(filePath, testCaseId1, currentUrl.Equals("http://localhost:49921/Admin/KhachHang") ? "Pass" : "Fail");
            // Kiểm tra xem khách hàng còn tồn tại trong cơ sở dữ liệu hay không và cập nhật kết quả
            VerifyCustomerDeleted();
            string testCaseId2 = "KH_02";
            UpdateTestResult(filePath, testCaseId2, !existsInDatabase ? "Pass" : "Fail");
        }

        private void DeleteCustomer()
        {
            // Xóa khách hàng
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/table/tbody/tr[2]/td[5]/a[3]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-default")).Click();
            Thread.Sleep(1000);
            // Đợi và kiểm tra URL
            wait.Until(ExpectedConditions.UrlToBe("http://localhost:49921/Admin/KhachHang"));
            currentUrl = driver.Url;
        }

        private void VerifyCustomerDeleted()
        {
            // Đọc dữ liệu test từ Excel
            DataTable testData = ReadTestData("C:\\BDCLPM\\DataTest_Nam.xlsx");
            DataRow testDataRow = testData.Rows[0]; // Nếu bạn chỉ có một hàng dữ liệu, sử dụng Rows[0]
            string ma_kh = testDataRow["ma_kh"].ToString();
            // Kiểm tra dữ liệu trong cơ sở dữ liệu
            existsInDatabase = CheckCustomerExists(ma_kh);
        }
        private bool CheckCustomerExists(string ma_kh)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                string query = "SELECT COUNT(1) FROM tblKhachHang WHERE ma_kh = @ma_kh";
                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ma_kh", ma_kh);
                    connection.Open();
                    int customerCount = (int)command.ExecuteScalar();
                    return customerCount > 0;
                }
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
