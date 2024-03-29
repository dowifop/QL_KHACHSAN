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
    internal class TestNhanPhong
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
            driver.FindElement(By.LinkText("Đặt Phòng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("#page-top > div > div.container-fluid > div > div:nth-child(1) > div > a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div/div[1]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[7]/a")).Click();
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
                    return result.Tables[4]; // Lấy sheet đầu tiên trong file Excel
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
        public void RoomWithTestData()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            int testCaseIndex = 1;
            foreach (DataRow row in testData.Rows)
            {             
                string testCaseId = $"NP_0{testCaseIndex}";
                string hoten1 = row["hoten1"].ToString();
                string tuoi1 = row["tuoi1"].ToString();
                string hoten2 = row["hoten2"].ToString();
                string tuoi2 = row["tuoi2"].ToString();
                string hoten3 = row["hoten3"].ToString();
                string tuoi3 = row["tuoi3"].ToString();
                string hoten4 = row["hoten4"].ToString();
                string tuoi4 = row["tuoi4"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                try
                {
                    // Điền thông tin vào form đăng ký
                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("hoten1")));
                    driver.FindElement(By.Id("hoten1")).Clear();
                    driver.FindElement(By.Id("hoten1")).SendKeys(hoten1);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("tuoi1")).Clear();
                    driver.FindElement(By.Id("tuoi1")).SendKeys(tuoi1);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("hoten2")).Clear();
                    driver.FindElement(By.Id("hoten2")).SendKeys(hoten2);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("tuoi2")).Clear();
                    driver.FindElement(By.Id("tuoi2")).SendKeys(tuoi2);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("hoten3")).Clear();
                    driver.FindElement(By.Id("hoten3")).SendKeys(hoten3);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("tuoi3")).Clear();
                    driver.FindElement(By.Id("tuoi3")).SendKeys(tuoi3);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("hoten4")).Clear();
                    driver.FindElement(By.Id("hoten4")).SendKeys(hoten4);
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("tuoi4")).Clear();
                    driver.FindElement(By.Id("tuoi4")).SendKeys(tuoi4);
                    Thread.Sleep(1000);
                    
                    driver.FindElement(By.CssSelector("#page-top > div.content-wrapper > div > div > form > div.form-group > div > input")).Click();
                    Thread.Sleep(2000); 

                    string newUrl = driver.Url;

                    if (testCaseId == "NP_09" && newUrl.Contains("/Result"))
                    {
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Pass");
                    }
                    else
                    {
                        UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                        driver.Navigate().GoToUrl("http://localhost:49921/Admin/HoaDon/Add/25");
                    }


                }
                catch (Exception ex)
                {
                  
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Fail");
                 
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
