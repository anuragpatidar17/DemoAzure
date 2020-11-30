using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using NPOI.SS.UserModel;
using System.Threading;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Reflection;



namespace SeleniumDemo
{
    
    public class Class1
    {
        [Test]
        public void SearchForWord()
        {
            var driver = new ChromeDriver();

            {
                //Notice navigation is slightly different than the Java version
                //This is because 'get' is a keyword in C#
                driver.Navigate().GoToUrl("http://www.google.com/");

                // Find the text input element by its name
                IWebElement query = driver.FindElement(By.Name("q"));
                
                //Excel
                FileStream file = new FileStream("D:\\a\\1\\s\\MDS.xlsx", FileMode.Open, FileAccess.Read);
                XSSFWorkbook workbook = new XSSFWorkbook(file1);
                ISheet sheet = workbook.GetSheet("Sheet");

                var value =string.Format(sheet.GetRow(1).GetCell(0).StringCellValue)
                
                // Enter something to search for
                query.SendKeys(value);
                console.WriteLine(value);

                // Now submit the form. WebDriver will find the form for us from the element
                query.Submit();

                // Google's search is rendered dynamically with JavaScript.
                // Wait for the page to load, timeout after 10 seconds
              

                driver.Quit();
            }

        }
    }
}
