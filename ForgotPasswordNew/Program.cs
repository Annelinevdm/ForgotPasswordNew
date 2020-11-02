using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.IO;
using System.Linq;

namespace ForgotPasswordNew
{
    public class Program
    {
        static void Main(string[] args)
        {
            //Open Chrome Driver
            IWebDriver driver = new ChromeDriver(@"C:\Users\27828\source\repos\");
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //Path to your excel file
            string path = "C:\\DATA\\ForgotPassword.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            //Get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            //Initialise variables
            int i = 2;
            int j = 1;

            //Open Environment
            string URl = worksheet.Cells[i, j].Value.ToString();
            driver.Navigate().GoToUrl((string)URl);
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            while (i <= rows)
            {
                //Click on the Link Forgot your password?
                IWebElement ForgotPasswordLink = driver.FindElement(By.Id("forgot-password"));
                ForgotPasswordLink.Click();

                //Enter Email
                string EMailAddress = worksheet.Cells[i, 2].Value.ToString();
                IWebElement EmailAddress = driver.FindElement(By.Id("Input_Email"));
                EmailAddress.SendKeys((string)EMailAddress);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                //Click on the Reset Password button
                Actions action1 = new Actions(driver);
                action1.MoveToElement(driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[2]/form/div[5]/div/button"))).Build().Perform();
                IWebElement ResetButton = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[2]/form/div[5]/div/button"));
                ResetButton.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                //Click on the Login Link
                IWebElement LoginLink = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div[2]/div/div[1]/a"));
                LoginLink.Click();

                i = i + 1;

            }
            //Quit Browser
            driver.Quit();
            package.Dispose();
        }
    }
}
