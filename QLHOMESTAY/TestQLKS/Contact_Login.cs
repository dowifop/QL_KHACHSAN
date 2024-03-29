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
using Microsoft.Data.SqlClient;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace TestQLKS
{
    public class Contact_Login
    {

        [TestFixture]
        public class ContactLoginTest
        {
            private IWebDriver driver;
            private WebDriverWait wait;

            [SetUp]
            public void SetUp()
            {
                // Register the code page provider to ensure encoding 1252 is available
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                driver = new ChromeDriver();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Adjust the time as necessary.
                driver.Navigate().GoToUrl("http://localhost:49921/Home/Contact");
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
                var worksheet = workbook.Worksheet(1); // Giả sử kết quả test nằm ở sheet đầu tiên

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
            public void contactLogin()
            {

                // Đọc dữ liệu test từ file Excel
                var testData = ReadTestData("C:\\Users\\TIEN\\Documents\\DBCL\\DataTestWeb.xlsx");
                int testCaseIndex = 1;
                foreach (DataRow row in testData.Rows)
                {
                    // Define testCaseId at the beginning of the loop
                    string testCaseId = $"PHANHOI_{testCaseIndex}";

                    // Lấy thông tin từ datatest
                    string username = row["UserName"].ToString();
                    string email = row["Email"].ToString();
                    string message = row["TestMessenger"].ToString();
                    string rating = row["Rating"].ToString();
                    string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                  
                    int ratingValue;
                    string actualErrorMessage = "";
                    bool isDateSelectionSuccessful = true;

                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ho_ten")));
                        driver.FindElement(By.Id("ho_ten")).Clear();
                        driver.FindElement(By.Id("ho_ten")).Click();
                        driver.FindElement(By.Id("ho_ten")).SendKeys(username);

                        driver.FindElement(By.Id("mail")).Clear();
                        driver.FindElement(By.Id("mail")).SendKeys(email);
                        Thread.Sleep(1000);


                        driver.FindElement(By.Id("noi_dung")).Clear();
                        var noiDungInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("noi_dung")));
                        //noiDungInput.Click();
                        noiDungInput.SendKeys(message);
                        Thread.Sleep(1000);
                        
                        if (int.TryParse(rating, out ratingValue))
                        {
                            // Nếu chuyển đổi thành công, tính starIndex dựa trên ratingValue
                            int starIndex = 6 - ratingValue;
                            // Các bước tiếp theo của việc xử lý rating...
                            var RatingInput = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath($"html/body/div[3]/div/div/div/div[1]/form/div[4]/div/div/div/label[{starIndex}]")));
                          
                            RatingInput.Click();
                            Thread.Sleep(1000);
                        }
                        else
                        {  
                            ratingValue = 0; 
                        }

                        // Click on the submit button
                        driver.FindElement(By.CssSelector("input[type='submit']")).Click();
                        Thread.Sleep(1000);

                        try
                        {
                            IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
                            actualErrorMessage = alert.Text;
                            alert.Accept();
                            isDateSelectionSuccessful = actualErrorMessage.Contains(expectedErrorMessage);
                        }
                       
                        catch (UnhandledAlertException ex)
                        {
                            Console.WriteLine("There was an unhandled alert after clicking the submit button: " + ex.AlertText);
                            UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Fail");
                        }
                        catch (WebDriverTimeoutException ex)
                        {
                            Console.WriteLine("The alert did not appear within the specified time: " + ex.Message);
                            UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Fail");
                        }
                        UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, isDateSelectionSuccessful ? "Pass" : "Fail");
                        if (string.IsNullOrEmpty(expectedErrorMessage))
                        {
                            // Giả sử đã có phương thức VerifyDataInDatabase()
                            VerifyDataInDatabase(username, email, message, rating);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Test failed for test case ID: {testCaseId} with exception: {ex.Message}");
                        UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Fail");
                        isDateSelectionSuccessful = false;
                    }
                    testCaseIndex++;
                }
            }

            private void VerifyDataInDatabase(string ho_ten, string mail, string noi_dung, string danh_gia)
            {
                // Chuỗi kết nối cập nhật
                string connectionString = "data source=.;initial catalog=dataQLKS;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework";

                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    var query = "SELECT ho_ten, mail, noi_dung, danh_gia FROM tblKhachHang WHERE ho_ten = @ho_ten";
                    using (var command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@ho_ten", ho_ten);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                Assert.AreEqual(ho_ten, reader["ho_ten"].ToString(), "Họ tên không khớp");
                                Assert.AreEqual(mail, reader["mail"].ToString(), "mail không khớp");
                                Assert.AreEqual(noi_dung, reader["noi_dung"].ToString(), "noi dung không khớp");
                                Assert.AreEqual(danh_gia, reader["danh_gia"].ToString(), "danh gia không khớp");
                            }
                            else
                            {
                                Assert.Fail("Không tìm thấy dữ liệu trong database cho Khach hang: " + ho_ten);
                            }
                        }
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
}
