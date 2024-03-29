using ClosedXML.Excel;
using ExcelDataReader;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
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

namespace TestQLKS
{
    internal class SuaPhongTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        bool isRoomSuccessful = false;  

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
            driver.FindElement(By.LinkText("Tùy Chỉnh")).Click();
            {
                var element = driver.FindElement(By.TagName("body"));
                Actions builder = new Actions(driver);
                builder.MoveToElement(element, 0, 0).Perform();
            }
            driver.FindElement(By.LinkText("Phòng")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[1]/div[1]/div/div[2]/div/div/div[2]/div/table/tbody/tr[8]/td[5]/a[1]")).Click();
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
                    return result.Tables[3]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(4); // Giả sử kết quả test nằm ở sheet đầu tiên

            bool isTestCaseFound = false;

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                // Giả sử cột 'A' chứa ID của test case
                if (row.Cell("A").Value.ToString() == testCaseID)
                {
                    isTestCaseFound = true;

                    row.Cell("E").SetValue(result);
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
        public void TestEditRoom()
        {
            var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest_Nam.xlsx");
            int testCaseIndex = 1;

            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"S_0{testCaseIndex}";
                string so_phong = row["so_phong"].ToString();
                string loai_phong = $"//option[. = '{row["loai_phong"]}']"; 
                string ma_tang = $"//option[. = '{row["ma_tang"]}']"; 
                try
                {
                    var so_phongElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("so_phong")));
                    so_phongElement.Click();
                    so_phongElement.Clear();
                    so_phongElement.SendKeys(so_phong);

                    var loai_phongElement = wait.Until(ExpectedConditions.ElementExists(By.Id("loai_phong")));
                    loai_phongElement.FindElement(By.XPath(loai_phong)).Click();

                    var ma_tangElement = wait.Until(ExpectedConditions.ElementExists(By.Id("ma_tang")));
                    ma_tangElement.FindElement(By.XPath(ma_tang)).Click();

                    var submitButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#page-top > div.content-wrapper > div.container-fluid > div.card.mb-3 > div.card-body > form > div > div:nth-child(6) > div > input")));
                    submitButton.Click();

                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/Admin/Phong"));

                    // Check the current URL to determine if the room edit was apparently successful
                    isRoomSuccessful = driver.Url.Contains("http://localhost:49921/Admin/Phong");

                    // Now check if the room actually exists in the database
                    if (isRoomSuccessful)
                    {
                        int loai_phong_value = Convert.ToInt32(row["loai_phong"]);
                        int ma_tang_value = Convert.ToInt32(row["ma_tang"]); 

                        // Call the method to check the database
                        isRoomSuccessful = DoesRoomExistInDatabase(so_phong, loai_phong_value, ma_tang_value);
                    }
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase_Nam.xlsx", testCaseId, isRoomSuccessful ? "Pass" : "Fail");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with error: {ex.Message}");
                    UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase_Nam.xlsx", testCaseId, "Fail");
                    isRoomSuccessful = false;
                }
               
            }
        }

        private bool DoesRoomExistInDatabase(string so_phong, int loai_phong, int ma_tang)
        {
            string connectionString = "data source=.;initial catalog=dataQLKS;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework"; 
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT COUNT(1) FROM tblPhong WHERE so_phong = @so_phong AND loai_phong = @loai_phong AND ma_tang = @ma_tang";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@so_phong", so_phong);
                command.Parameters.AddWithValue("@loai_phong", loai_phong);
                command.Parameters.AddWithValue("@ma_tang", ma_tang);

                connection.Open();
                int count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
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
