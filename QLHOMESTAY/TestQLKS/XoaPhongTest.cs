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
    internal class XoaPhongTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        bool isRoomSuccessful = false;

        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.
                CodePagesEncodingProvider.Instance);
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
          
        }
  

        private void UpdateTestResult(string filePath, string testCaseID, 
            string result)
        {
            // Cập nhật kết quả test trong file Excel
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(4); 
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
        public void TestDeleteRoom()
        {
            driver.FindElement(By.XPath("/html/body/div[1]" +
                "/div[1]/div/div[2]/div/div/div[2]/div/table" +
                "/tbody/tr[1]/td[5]/a[2]")).Click();
            Thread.Sleep(1000);

            driver.FindElement(By.CssSelector("#page-top > " +
                "div.content-wrapper > div.container-fluid > " +
                "div > div.card-body > div > form > div > input")).Click();
            Thread.Sleep(1000);

            wait = new WebDriverWait(driver, TimeSpan.
                FromSeconds(10));
            wait.Until(ExpectedConditions.
                UrlContains("http://localhost:49921/Admin/Phong"));
            isRoomSuccessful = driver.Url.
                Contains("http://localhost:49921/Admin/Phong");
            string so_phong = "P101"; 
            isRoomSuccessful = !DoesRoomExistInDatabase(so_phong); 

            string testCaseID = "X_01"; 
            UpdateTestResult("C:\\Users" +
                "\\dowif\\Documents\\DBCLPM\\" +
                "Testcase_Nam.xlsx", testCaseID, 
                isRoomSuccessful ? "Pass" : "Fail");
        }
        private bool DoesRoomExistInDatabase(string so_phong)
        {
            string connectionString = "data source=.;" +
                "initial catalog=dataQLKS;integrated security=True;" +
                "trustservercertificate=True;MultipleActiveResultSets=" +
                "True;App=EntityFramework"; 
            using (SqlConnection connection = new 
                SqlConnection(connectionString))
            {
                string query = "SELECT COUNT(1) " +
                    "FROM tblPhong " +
                    "WHERE so_phong = @so_phong";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@so_phong", so_phong);

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

