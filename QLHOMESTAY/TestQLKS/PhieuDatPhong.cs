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
using System.Data.SqlClient;

namespace TestQLKS
{
    internal class PhieuDatPhong
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private string currentUrl;
        private bool existsInDatabase;
        private string connectionString = "data source=.;initial catalog=dataQLKS;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework";
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
            driver.FindElement(By.Id("ma_kh")).SendKeys("binh");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("mat_khau")).SendKeys("123456");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/a")).Click();
            Thread.Sleep(1000);
        }

        [Test]
        public void DeleteAndVerifyRent()
        {


            // Xóa khách hàng và kiểm tra
            CancelRent();
            string testCaseId1 = "1";
            string filePath = "C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx";
            UpdateTestResult(filePath, testCaseId1, currentUrl.Equals("http://localhost:49921/Home/BookRoom") ? "Pass" : "Fail");
            // Kiểm tra xem khách hàng còn tồn tại trong cơ sở dữ liệu hay không và cập nhật kết quả
            VerifyRentCanceled();
            string testCaseId2 = "2";
            UpdateTestResult(filePath, testCaseId2, !existsInDatabase ? "Pass" : "Fail");
        }
     
        public void CancelRent()
        {
            // Đọc dữ liệu test từ file Excel
            //var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
           
            driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[6]/a[1]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/form/div/input")).Click();
            Thread.Sleep(1000);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/Home/BookRoom")); 
                Assert.That(driver.Url, Does.Contain("http://localhost:49921/Home/BookRoom"));
            currentUrl = driver.Url;

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
                    return result.Tables[4];
                }
            }
        }
        private void VerifyRentCanceled()
        {
            // Đọc dữ liệu test từ Excel
            DataTable testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            DataRow testDataRow = testData.Rows[0]; // Nếu bạn chỉ có một hàng dữ liệu, sử dụng Rows[0]
            string ma_kh = testDataRow["ma_kh"].ToString();
            // Kiểm tra dữ liệu trong cơ sở dữ liệu
            existsInDatabase = CheckRentExists(ma_kh);
        }
        private bool CheckRentExists(string ma_kh)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                string query = @"
            SELECT COUNT(*)
            FROM tblPhieuDatPhong
            WHERE ma_kh = @ma_kh AND ma_tinh_trang = 3"; 

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ma_kh", ma_kh);
                    connection.Open();
                    int rentCount = (int)command.ExecuteScalar();
                    return rentCount == 0;
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(5); 

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
