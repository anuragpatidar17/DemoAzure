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
using obj = DemoAzure.CommonFunctions;
using NPOI.SS.Util;
using NPOI.HSSF.Record;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;

namespace DemoAzure
{
	

	public class Form_5100_maintenanceIssue
	{

		static String form_name = "maintenanceIssue";
		static string value, radio, comments_path, submit;
		static IWebDriver driver;
		static XSSFWorkbook workbook;
		static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
		static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
		//get current date time with Date()
		static DateTime date = DateTime.Now;
		// Now format the date
		static String date1 = dateFormat.Format(date);


		[Test]
		public static void Test()
		{
			FileStream file1 = new FileStream(@"D:\a\1\s\DemoAzure\MDS.xlsx", FileMode.Open, FileAccess.Read);
			XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
			ISheet sheet1 = workbook1.GetSheet("Sheet1");
			Console.WriteLine(date1);
			// Local Selenium WebDriver
			value = obj.F_ReadFromExcel(file1, workbook1, sheet1, 1, 0);
			driver = new ChromeDriver();

			//open url
			value = obj.F_ReadFromExcel(file1, workbook1, sheet1, 3, 0);
			obj.F_OpenUrl(driver, value);
			Thread.Sleep(25000);


			//Navigations:
			//Click on Environmental tab
			obj.F_ClickBy_linkText(driver, "Environmental");

			Thread.Sleep(1000);

			//click on Maintenance Issue Notification
			obj.F_ClickBy_linkText(driver, "Maintenance Issue Notification");
			Thread.Sleep(15000);

			//excel

			FileStream file = new FileStream(@"D:\a\1\s\DemoAzure\MDS.xlsx", FileMode.Open, FileAccess.Read);
			workbook = new XSSFWorkbook(file);
			ISheet sheet = workbook.GetSheet(form_name);



			//Department field:
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 0);
			obj.F_Sendkeys_name(driver, "department", value);
			Thread.Sleep(1000);

			//enter value in FIM/HAZobj field:

			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 1);
			obj.F_Sendkeys_name(driver, "synergiNumber", value);
			Thread.Sleep(1000);


			//Enter contact no.
			string contact = obj.F_ReadFromExcel(file, workbook, sheet, 1, 2);


			Thread.Sleep(1000);
			if (value != "")
			{
				obj.F_Sendkeys_name(driver, "contactNo", contact);
			}
			else
			{
				Assert.Fail("Value selected in contact No. field is null");
			}
			//Enter contact email

			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 3);

			if (value != "")
			{
				obj.F_Sendkeys_name(driver, "contactEmail", value);
				Console.Write("Email found");
				Thread.Sleep(1000);
			}
			else
			{
				Assert.Fail("Value selected in contact Email field is" + "\t" + null);
			}

			//if (value == "")
			//




			//Select value for Infrastructure owner
			String Infra_owner = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(1) > section:nth-child(2) > label > select";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 4);
			obj.F_Sendkeys_Selector(driver, Infra_owner, value);

			Thread.Sleep(1000);
			//Select value for Infrastructure Type
			String Infrastructure_Type = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(1) > section:nth-child(4) > label > select";
			String Infra_type = obj.F_ReadFromExcel(file, workbook, sheet, 1, 5);
			obj.F_Sendkeys_Selector(driver, Infrastructure_Type, Infra_type);


			Thread.Sleep(1000);
			//Select the value in Issue
			String issue = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(2) > section:nth-child(2) > select";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 6);

			if (value != "")
			{
				obj.F_Sendkeys_Selector(driver, issue, value);
				Console.Write("Issue found");
				Thread.Sleep(1000);
			}
			else
			{
				Assert.Fail("Value selected in Issue field is  null");
			}


			Thread.Sleep(1000);
			//Condtion for field value::

			if (Infra_type.Equals("Facilities") == false && Infra_type.Equals("Ponds") == false)
			{
				if (Infra_type.Equals("Pipelines") == false)
				{
					String field = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(2) > section:nth-child(4) > label > select";
					value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 7);
					obj.F_Sendkeys_Selector(driver, field, value);
				}
				else
				{
					String field = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(2) > section:nth-child(4) > label > dropdown-with-add-item > div > table > tr:nth-child(1) > td:nth-child(1) > label > select";
					value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 7);
					obj.F_Sendkeys_Selector(driver, field, value);
				}
			}

			Thread.Sleep(1000);
			//Enter Infrastructure ID:
			String Infra_ID = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(3) > section:nth-child(2) > label > select";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 8);
			obj.F_Sendkeys_Selector(driver, Infra_ID, value);

			Thread.Sleep(1000);
			//enter value in lot of plan
			String lot = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(3) > section:nth-child(4) > label > input";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 9);


			if (value != "")
			{
				obj.F_Sendkeys_Selector(driver, lot, value);
				Console.Write("Lot on plan found");
				Thread.Sleep(1000);
			}
			else
			{
				Assert.Fail("Value selected in lot field is  null");
			}


			Thread.Sleep(1000);
			//enter value in KP field
			String KP = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(4) > section:nth-child(2) > label > input";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 10);
			obj.F_Sendkeys_Selector(driver, KP, value);

			Thread.Sleep(1000);
			//enter value in Priority Rating(guide)
			String priority = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(4) > section:nth-child(4) > select";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 11);
			obj.F_Sendkeys_Selector(driver, priority, value);

			Thread.Sleep(1000);
			//GPS locations::
			String gps_icon = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(5) > gps-coordinates > div > div > div > div:nth-child(1) > section:nth-child(2) > label > i > span";
			obj.F_ClickBy_selector(driver, gps_icon);

			Thread.Sleep(1000);



			//enter type value in distance required
			String distance = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(6) > section.col.col-4 > label > input";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 12);


			Thread.Sleep(1000);
			if (value != "")
			{
				obj.F_Sendkeys_Selector(driver, distance, value);
				Console.Write("distance found");
				Thread.Sleep(1000);
			}
			else
			{
				Assert.Fail("Value selected in distance field is  null");
			}

			//Is this Civil Maintenance Entry a Risk Level 3 Asset Integrity Risk or Higher ?
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 13);
			radio = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(7) > section:nth-child(2) > div > label:nth-child(" + value[0] + ") > i";
			obj.F_ClickBy_selector(driver, radio);


			Thread.Sleep(1000);
			//Is this Civil Maintenance Entry a Regulatory Notifiable Issue ?
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 14);
			radio = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(8) > section:nth-child(2) > div > label:nth-child(" + value[0] + ") > i";
			obj.F_ClickBy_selector(driver, radio);
			Thread.Sleep(1000);

			if (value[0] == '1')
			{
				comments_path = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(8) > section:nth-child(3) > label > input";
				value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 15);
				obj.F_Sendkeys_Selector(driver, comments_path, value);
			}

			Thread.Sleep(1000);

			//Is this Civil Maintenance entry a formal landholder complaint ?
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 16);
			radio = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(9) > section:nth-child(2) > div > label:nth-child(" + value[0] + ") > i";
			obj.F_ClickBy_selector(driver, radio);
			if (value[0] == '1')
			{
				comments_path = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(9) > section:nth-child(3) > label > input";
				value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 17);
				obj.F_Sendkeys_Selector(driver, comments_path, value);
			}

			Thread.Sleep(1000);

			//Does Maintenance Entry Require on-ground Validation ?
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 18);
			radio = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(10) > section:nth-child(2) > div > label:nth-child(" + value[0] + ") > i";
			obj.F_ClickBy_selector(driver, radio);

			Thread.Sleep(1000);
			if (value[0] == '1')
			{
				comments_path = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(10) > section.col.col-4 > label > input";
				value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 19);
				obj.F_Sendkeys_Selector(driver, comments_path, value);
			}

			Thread.Sleep(1000);
			//Maintenance Issue Pictures
			//edit pictures
			String edit_pictures = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(11) > section.col.col-4 > button-edit-pictures > button";

			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 20);
			if (value.Equals("Yes"))

			{
				obj.F_ClickBy_selector(driver, edit_pictures);
				int n = Int32.Parse(obj.F_ReadFromExcel(file, workbook, sheet, 1, 21));
				int c = 22;   //21+1
				int r = 1;
				value = obj.F_ReadFromExcel(file, workbook, sheet, 1, c);

				String file_mode = "body > modal-container > div > div > div.modal-body > edit-pictures > div > table:nth-child(2) > tr:nth-child(1) > td:nth-child(4) > div > label";

				Thread.Sleep(4000);

				obj.F_ClickBy_selector(driver, file_mode);
				Thread.Sleep(2000);

				for (int i = 1; i <= n; i++)
				{

					driver.FindElement(By.Id("myFileLookup")).SendKeys(value);  //img upload
					Thread.Sleep(1000);
					r++;
					if (i < n)
					{
						String add_media = "body > modal-container > div > div > div.modal-body > edit-pictures > div > table:nth-child(2) > tr:nth-child(2) > td:nth-child(3) > button.btn.btn-success";
						Thread.Sleep(1000);
						obj.F_ClickBy_selector(driver, add_media);

						value = obj.F_ReadFromExcel(file, workbook, sheet, r, c);
						Thread.Sleep(1000);
					}
				}
				String close = "body > modal-container > div > div > div.modal-header.alert-success > table > tr > td:nth-child(2) > i";
				obj.F_ClickBy_selector(driver, close);
			}

			Thread.Sleep(1000);

			//Enter description:
			String description = "#wid-id-1 > div > div > form > fieldset:nth-child(3) > civil-maintenance > div > section > div:nth-child(13) > section > label > textarea";
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 23);

			if (value != "")
			{
				obj.F_Sendkeys_Selector(driver, description, value);
				Console.Write("description found");
				Thread.Sleep(1000);
			}
			else
			{
				Assert.Fail("Value selected in Description field is  null");
			}


			Thread.Sleep(1000);
			//Comments
			value = obj.F_ReadFromExcel(file, workbook, sheet, 1, 24);
			comments_path = "comment";
			obj.F_Sendkeys_name(driver, comments_path, value);

			Thread.Sleep(2000);


			//Submit
			submit = "#wid-id-1 > div > div > form > footer > div > section > button.btn.btn-primary";

			if (obj.F_FindElement_selector(driver, submit).Enabled)
			{
				obj.F_ClickBy_selector(driver, submit);
			}

			else
			{

				((IJavaScriptExecutor)driver).ExecuteScript("alert('Submit button not enabled')");
				Console.WriteLine("Submit button not enabled");
				Assert.Fail("Submit button not enabled:");

			}
			Thread.Sleep(3000);

			//confirm
			obj.F_ClickBy_selector(driver, "body > modal-container > div > div > app-check-save-modal > div.modal-footer > button:nth-child(1)");

			//wait until report is clickable:
			Thread.Sleep(4000);




			string operatorName = obj.F_FindElement_selector(driver, "#left-panel > sa-login-info > div > span > a > span").Text;
			Console.WriteLine(operatorName);

			//Reports button

			obj.F_ClickBy_linkText(driver, "Reports");

			Thread.Sleep(10000);

			//Select Maintenance in function Dropdown
			String Function = "//*[@id=\"repeatSelect\"]";
			driver.FindElement(By.XPath(Function)).SendKeys("Environmental");
			Thread.Sleep(1000);

			//Select drivehead in report dropdown
			string Report = "/html/body/app-root/app-main-layout/div/div/div/div/app-report/div/div/article/div/div/div/form/fieldset[2]/div/section[4]/label/select";
			driver.FindElement(By.XPath(Report)).SendKeys("Maintenance Issue Notification");

			//Click Submit
			Thread.Sleep(1000);

			obj.F_ClickBy_selector(driver, "#wid-id-1 > div > div > form > fieldset:nth-child(2) > div:nth-child(1) > section:nth-child(5) > button");

			Thread.Sleep(1000);

			//Click Export to Excel
			obj.F_ClickBy_selector(driver, "#wid-id-1 > div > div > form > fieldset:nth-child(3) > div:nth-child(4) > section > kendo-grid > kendo-grid-toolbar > button");


			Thread.Sleep(3000);
			string userRoot = System.Environment.GetEnvironmentVariable("USERPROFILE");
			string downloadFolder = Path.Combine(userRoot, "Downloads");
			Console.WriteLine(downloadFolder);
			//Open excel report
			FileStream file2 = new FileStream(downloadFolder + "\\Export.xlsx", FileMode.Open, FileAccess.Read);
			XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
			ISheet sheet2 = workbook2.GetSheet("Sheet1");



			if (obj.F_ReadFromExcel(file2, workbook2, sheet2, 1, 1).Equals(date1, StringComparison.InvariantCultureIgnoreCase) && obj.F_ReadFromExcel(file2, workbook2, sheet2, 1, 4).Equals(operatorName, StringComparison.InvariantCultureIgnoreCase) && obj.F_ReadFromExcel(file2, workbook2, sheet2, 1, 5).Equals(contact, StringComparison.InvariantCultureIgnoreCase))
			{
				((IJavaScriptExecutor)driver).ExecuteScript("alert('EXCEL REPORT VERIFIED')");
				Console.WriteLine("EXCEL REPORT VERIFIED");
			}
			else
				Console.WriteLine("Excel Report not verified");

			//file2.Close();

			string fileName = "Export.xlsx";
			string sourcePath = downloadFolder;
			string targetFile = DateTime.Now.ToString("MM-dd-yy-hh-mm-ss-") + fileName;
			string targetPath = desktopPath + "//MDSTestData//" + form_name;

			string sourceFile = Path.Combine(sourcePath, fileName);

			if (!Directory.Exists(targetPath))
			{
				Directory.CreateDirectory(targetPath);
			}

			string destFile = targetPath + "//" + targetFile;

			File.Move(sourceFile, destFile);

			driver.Quit();




		}



	}

}