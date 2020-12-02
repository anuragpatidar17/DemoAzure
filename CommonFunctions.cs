using NUnit.Framework;
using OpenQA.Selenium;
using System.Linq;
using System.Text;
using System.Reflection;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using NPOI.SS.UserModel;
using System.Threading;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DemoAzure
{
	class CommonFunctions
	{






		public static void F_OpenUrl(IWebDriver driver, String URL)
		{


			driver.Navigate().GoToUrl(URL);
			driver.Manage().Window.Maximize();
			driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
		}



		public static void F_Sendkeys_Selector(IWebDriver driver, String cssSelector, String value)
		{
			driver.FindElement(By.CssSelector(cssSelector)).SendKeys(value);
		}

		public static void F_Sendkeys_name(IWebDriver driver, String name, String value)
		{
			driver.FindElement(By.Name(name)).SendKeys(value);
		}

		public static void F_Sendkeys_class(IWebDriver driver, String className, String value)
		{
			driver.FindElement(By.ClassName(className)).SendKeys(value);
		}



		public static void F_ClickBy_selector(IWebDriver driver, String cssSelector)
		{
			driver.FindElement(By.CssSelector(cssSelector)).Click();
		}

		public static void F_ClickBy_name(IWebDriver driver, String name)
		{
			driver.FindElement(By.Name(name)).Click();
		}

		public static void F_ClickBy_className(IWebDriver driver, String className)
		{
			driver.FindElement(By.ClassName(className)).Click();
		}

		public static void F_ClickBy_linkText(IWebDriver driver, String linkText)
		{
			driver.FindElement(By.LinkText(linkText)).Click();
		}


		public static void F_ValidateLink_linkText(IWebDriver driver, String linkText)
		{
			if (driver.FindElement(By.LinkText(linkText)).Size != null)
			{
				Console.WriteLine("Required Link Found");
			}
			else
				((IJavaScriptExecutor)driver).ExecuteScript("alert('Required link not found')");

		}



		public static String F_GetElementText(IWebDriver driver, String cssSelector)
		{
			String WellCode = driver.FindElement(By.CssSelector(cssSelector)).Text;
			return WellCode;
		}




		//kamal

		public static String F_ReadFromExcel(FileStream file, XSSFWorkbook workbook, ISheet sheet, int r, int c)   //r=row, c=col
		{

			return string.Format(sheet.GetRow(r).GetCell(c).StringCellValue);

		}


		public static IWebElement F_FindElement_selector(IWebDriver driver, String cssSelector)
		{
			IWebElement chk = driver.FindElement(By.CssSelector(cssSelector));
			return chk;
		}

	}
}
