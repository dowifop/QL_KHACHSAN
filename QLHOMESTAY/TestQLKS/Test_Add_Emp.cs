using ExcelDataReader;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using SeleniumExtras.WaitHelpers;

namespace TestQLKS
{
    public class Test_Add_Emp
    {
        [TestFixture]
        public class ContactLoginTest
        {
            private IWebDriver driver1;
            private WebDriverWait wait;

            [SetUp]
            public void SetUp()
            {
                // Register the code page provider to ensure encoding 1252 is available
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                driver1 = new ChromeDriver();
                wait = new WebDriverWait(driver1, TimeSpan.FromSeconds(10)); // Adjust the time as necessary.
                driver1.Navigate().GoToUrl("http://localhost:49921/Admin/Index/Login");
                Thread.Sleep(1000);
                driver1.FindElement(By.Id("tai_khoan")).Click();
                driver1.FindElement(By.Id("tai_khoan")).SendKeys("admin");
                driver1.FindElement(By.Id("mat_khau")).Click();
                driver1.FindElement(By.Id("mat_khau")).SendKeys("12345");
                driver1.FindElement(By.CssSelector(".btn")).Click();
                driver1.FindElement(By.CssSelector(".nav-item:nth-child(6) > .nav-link")).Click();
                Thread.Sleep(1000);
                driver1.FindElement(By.LinkText("Nhân viên")).Click();
                driver1.FindElement(By.LinkText("Thêm nhân viên")).Click();

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
                        return result.Tables[3]; // Lấy sheet thu 3 trong file Excel
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

            [TearDown]
            protected void TearDown()
            {
                driver1.Quit();
                driver1.Dispose();
            }

            [Test]
            public void Add_Emp()
            {
                // Đọc dữ liệu test từ file Excel
                var testData = ReadTestData("C:\\Users\\TIEN\\Documents\\DBCL\\DataTestWeb.xlsx");
                int testCaseIndex = 1;
                foreach (DataRow row in testData.Rows)
                {
                    // Define testCaseId at the beginning of the loop
                    string testCaseId = $"Add_NV_{testCaseIndex}";

                    // Lấy thông tin từ datatest
                    string username = row["Hoten"].ToString();
                    string phone = row["SDT"].ToString();
                    string account = row["Account"].ToString();
                    string password = row["Password"].ToString();
                    string Chucvu = row["Chucvu"].ToString();
                    //string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                    string actualErrorMessage = "";
                    bool isDateSelectionSuccessful = true;

                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ho_ten")));
                        driver1.FindElement(By.Id("ho_ten")).Clear();
                        driver1.FindElement(By.Id("ho_ten")).Click();
                        driver1.FindElement(By.Id("ho_ten")).SendKeys(username);

                        driver1.FindElement(By.Id("sdt")).Clear();
                        driver1.FindElement(By.Id("sdt")).SendKeys(phone);
                        Thread.Sleep(1000);


                        driver1.FindElement(By.Id("tai_khoan")).Clear();
                        driver1.FindElement(By.Id("tai_khoan")).SendKeys(account);
                        Thread.Sleep(1000);

                        driver1.FindElement(By.Id("mat_khau")).Clear();
                        driver1.FindElement(By.Id("mat_khau")).SendKeys(password);
                        Thread.Sleep(1000);


                        driver1.FindElement(By.Id("ma_chuc_vu")).Click();
                        
                            {
                            var dropdown = driver1.FindElement(By.Id("ma_chuc_vu"));
                            dropdown.FindElement(By.XPath($"//option[. = '{Chucvu}']")).Click();
                             }

                        driver1.FindElement(By.Id("ma_chuc_vu")).SendKeys(password);
                        Thread.Sleep(1000);


                        driver1.FindElement(By.CssSelector(".btn-default")).Click();
                        Thread.Sleep(1000);
                        // Đợi cho đến khi URL thay đổi hoặc thêm một điều kiện chờ cụ thể khác
                        wait.Until(d => d.Url.Equals("http://localhost:49921/Admin/NhanVien"));
                        // Kiểm tra URL hiện tại sau khi chuyển hướng
                        if (driver1.Url.Equals("http://localhost:49921/Admin/NhanVien"))
                        {
                            UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Passed");
                            // Nếu đúng là trang danh sách nhân viên, chuyển đến trang tạo mới nhân viên
                            driver1.Navigate().GoToUrl("http://localhost:49921/Admin/NhanVien/Create");
                        }
                                   
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Test failed for test case ID: {testCaseId} with exception: {ex.Message}");
                        UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Failed");
                        isDateSelectionSuccessful = false;
                    }
                    testCaseIndex++;
                }
            }

        }
    }
}
