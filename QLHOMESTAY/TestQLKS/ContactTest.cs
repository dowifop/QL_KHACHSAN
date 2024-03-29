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
using DocumentFormat.OpenXml.Spreadsheet;

namespace TestQLKS
{
    public class ContactTest
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
                //driver.Manage().Window.Maximize();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Adjust the time as necessary.
                driver.Navigate().GoToUrl("http://localhost:49921/Account/Login");
                Thread.Sleep(1000);
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ma_kh")));
                driver.FindElement(By.Id("ma_kh")).Clear();
                driver.FindElement(By.Id("ma_kh")).SendKeys("tien1234");
                Thread.Sleep(1500);
                driver.FindElement(By.Id("mat_khau")).Clear();
                driver.FindElement(By.Id("mat_khau")).SendKeys("123456");

                driver.FindElement(By.CssSelector(".btn-primary")).Click();

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
                // Sử dụng tên hoặc chỉ mục chính xác của sheet
                var worksheet = workbook.Worksheet(1);
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
            public void Contact()
            {
                // Đọc dữ liệu test từ file Excel
                var testData = ReadTestData("C:\\Users\\TIEN\\Documents\\DBCL\\DataTestWeb.xlsx");
                int testCaseIndex = 5;


                foreach (DataRow row in testData.Rows)
                {
                    // Define testCaseId at the beginning of the loop
                    string testCaseId = $"PHANHOI_{testCaseIndex}";

                    // Lấy thông tin từ datatest
                    // string url = row["LinkTest"].ToString() ;
                    string message = row["Messtext"].ToString();
                    string rating = row["Rating"].ToString();
                    string erromessage = row["ExpectedErrorMessage"].ToString();

                    // code mới 
                   
                    int ratingValue;
                    string actualErrorMessage = "";
                    bool isRatingSuccessful = true;
                    try
                    {

                        wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText("Phản Hồi"))).Click();

                        var noiDungInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("noi_dung")));
                        noiDungInput.Click();
                        noiDungInput.Clear();
                        noiDungInput.SendKeys(message);

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

                        Thread.Sleep(1000);
                       

                        Thread.Sleep(1000);
                        // driver.FindElement(By.CssSelector(cssSelector)).Click();
                        //driver.FindElement(By.CssSelector(".star:nth-child({starIndex})")).Click();

                        driver.FindElement(By.CssSelector("input[type='submit']")).Click();

                        IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
                        if (alert.Text.Contains(erromessage))
                        {
                            alert.Accept();
                            UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Pass");
                        }
                        else
                        {
                            alert.Dismiss();
                            UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Fail");
                        }

                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Test failed for test case ID: {testCaseId} with exception: {ex.Message}");
                        UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Fail");
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
}
