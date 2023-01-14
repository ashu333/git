import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import javax.swing.plaf.synth.SynthOptionPaneUI;

import org.apache.poi.hpsf.Array;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Selinium_First {

	public static void main(String[] args) throws Exception {

		Selinium_First obj = new Selinium_First();
		/*
		 * List<List<String>> data = obj.readExcel("/home/ashish/Downloads",
		 * "Attendance Info Tracker _ 5th Jan.xlsx", "Sheet17");
		 * 
		 * List<List<String>> data2 = obj.readExcel("/home/ashish/Downloads",
		 * "Att 5th Jan.xlsx", "Sheet2");
		 * 
		 * //List<String> SubjectList = new ArrayList<>(); for (int i = 1; i <
		 * data.size(); i++) { List<String> a = data.get(i); String Subject = ""; for
		 * (int j = 0; j < a.size(); j++) { boolean repeat = false; /* if ((j == 1 || j
		 * == 6) && i != 0) { Subject += a.get(j) + "_"; } else if (j == 8 && i != 0) {
		 * Subject += a.get(j); }else if(j == 4 || j == 5 && i != 0) { Subject
		 * +=","+a.get(j); } else if (j == a.size() - 1) { Subject += "," + a.get(j); }
		 *
		 * if (i != 0) { Subject = a.get(8).trim() + "_" + a.get(1).trim() + "_" +
		 * a.get(6).trim() + "-" + a.get(6).trim() + "-" + a.get(4).trim() + "-" +
		 * a.get(5).trim() + "-" + a.get(a.size() - 1).trim(); break; }
		 */
		// System.out.print(a.get(j) + "|| ");
		/*
		 * if(!(a.get(2).equalsIgnoreCase("Check Out Time"))) { String employeeName =
		 * a.get(2).trim(); String first_namearr[] = employeeName.split(" "); String
		 * firstName = first_namearr[0].trim();
		 * 
		 * boolean onLeave = a.get(a.size()-1).equalsIgnoreCase("leave") ||
		 * a.get(a.size()-1).equalsIgnoreCase("week off") ? false:true; if(onLeave) {
		 * for (int k = 1; k < data2.size(); k ++) { List<String> l = data2.get(k);
		 * String timing = "";
		 * 
		 * if(k==65) { System.out.print(""); }
		 * 
		 * 
		 * if(l.get(2).contains(employeeName)) { System.out.println(employeeName);
		 * System.out.println(l.get(l.size()-1)); repeat =true; } }
		 * 
		 * }
		 * 
		 * } if (repeat) { break; } }
		 */
		/*
		 * System.out.println(); System.out.println(Subject); SubjectList.add(Subject);
		 * System.out.println();
		 */

		System.out.println("**********************************");
		// System.out.println("row number = "+i);
		// }

		// for attendence tracker ..

		System.setProperty("webdriver.chrome.driver",
				"/home/ashish/Downloads/selenium/driver/chromedriver_linux64/chromedriver");
		WebDriver driver = new ChromeDriver();
		String url = "https://supporthub.embee.co.in/a/tickets/view/new_and_my_open?order_by=created_at&order_type=desc&query_hash=%5B%7B%22value%22%3A%5B%2210%22%2C%222%22%5D%2C%22condition%22%3A%22status%22%2C%22operator%22%3A%22is_in%22%2C%22type%22%3A%22default%22%7D%2C%7B%22value%22%3A%22six_months%22%2C%22condition%22%3A%22created_at%22%2C%22operator%22%3A%22is_greater_than%22%2C%22type%22%3A%22default%22%7D%5D";
		System.out.println("hittimg url..");

		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofMinutes(1));
		driver.get(url);
		driver.manage().timeouts().implicitlyWait(Duration.ofMinutes(1));
		// driver.findElement(By.xpath("//a[@href='/support/login']")).click();
		// driver.manage().timeouts().implicitlyWait(Duration.ofHours(10));
		driver.findElement(By.xpath("//input[@id='username']")).sendKeys("support.nsd@embee.co.in");
		driver.findElement(By.xpath("//input[@id='password']")).sendKeys("Embee@1234#"); // Thread.sleep(10000); //
		driver.manage().timeouts().implicitlyWait(Duration.ofMinutes(1));
		driver.findElement(By.xpath("//button[@data-testid='login-button']")).click();
		Thread.sleep(20000);
	    obj.autoGrab(driver);	
	  //*[@id="ember783"]/div/div[3]/div/ul/li[16]

		/*
		 * }
		 * 
		 * for (int i = 16; i < SubjectList.size(); i++) {
		 * System.out.println("row number = " + i); obj.callEmbee(driver, url,
		 * SubjectList.get(i), i, obj); }
		 * 
		 * // obj.writeExcel("/home/ashish/Downloads", "Dec-2022 -Call details.xlsx", //
		 * "Sheet2", "307017");
		 * 
		 */
	}

	public void callEmbee(WebDriver driver, String url, String Subject, int rowNumber, Selinium_First obj)
			throws Exception {
		String arr[] = Subject.split("-");
		System.out.println("Subject = " + arr[0]);
		// String descarr[] = arr[0].split("_");
		String desc = arr[1];
		System.out.println("description =" + desc);
		String category = arr[2];
		System.out.println("Category = " + category);
		String subCategory = arr[3];
		System.out.println("SubCategory = " + subCategory);
		String remarks = arr[arr.length - 1];
		System.out.println("remarks = " + remarks);

		boolean formfilled = true;
		boolean ticketUpdate = true;
		try {
			driver.findElement(By.xpath("//*[@id='add-new-icon']")).click();
			// Thread.sleep(10000);
			driver.findElement(By.xpath("//*[@id='top_nav_new_list']/div[1]/span[1]/a/div[2]")).click();
		} catch (Exception e) {
			System.out.println("interrupted due to " + e);
			Thread.sleep(5000);
			driver.findElement(By.xpath("//*[@id='add-new-icon']")).click();
			// Thread.sleep(10000);
			driver.findElement(By.xpath("//*[@id='top_nav_new_list']/div[1]/span[1]/a/div[2]")).click();
		}
		while (formfilled) {
			try {
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_email']"))
						.sendKeys("MSD.DeepakNitrate@embee.co.in");
				driver.findElement(By.xpath("//html")).click();
				/*
				 * driver.findElement(By.xpath("//*[@id='select2-chosen-9']")).click();
				 * //System.out.println(SubjectList.size());
				 * 
				 * driver.findElement(By.xpath("//*[@id='helpdesk_ticket_subject']")).sendKeys(
				 * Subject); Thread.sleep(5000);
				 * //System.out.println(driver.findElement(By.xpath(
				 * "//*[@id='select2-result-label-37']/span")).getText());
				 * driver.findElement(By.xpath("//*[@id='s2id_autogen9_search']")).
				 * sendKeys("Embee-MSD Managed"); System.out.println("message passed ...");
				 * 
				 * element.sendKeys(Keys.TAB); System.out.println("tab pressed");
				 * 
				 * driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				 */
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_ticket_mode_348146']")).click();
				driver.findElement(By.xpath("//*[@id='select2-results-10']/li[2]")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_nsd_member_name_348146']"))
						.click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen13_search']")).sendKeys("pou");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_source']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen28_search']")).sendKeys("ema");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_status']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen29_search']")).sendKeys("assign ");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_group_id']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen33_search']")).sendKeys("mwp s");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_responder_id']")).click();
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen34_search']")).sendKeys("re");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_on_roaster_engineer_348146']"))
						.click();
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen15_search']")).sendKeys("rohi");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_department_id']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen35_search']")).sendKeys("deepak ni");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='NewTicket']/ul/li[23]/div/div[1]/div[2]")).sendKeys(desc);

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_category']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen16_search']")).sendKeys(category);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_sub_category']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen17_search']")).sendKeys(subCategory);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_location_348146']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen18_search']")).sendKeys("india-w");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(5000);

				System.out.println(Subject + "first");
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_subject']")).sendKeys(arr[0]);
				System.out.println(Subject);

				Thread.sleep(4000);
				driver.findElement(By.xpath("//*[@id='select2-chosen-9']")).click();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen9_search']")).sendKeys("Embee-MSD Managed");
				System.out.println("message passed ...");
				/*
				 * element.sendKeys(Keys.TAB); System.out.println("tab pressed");
				 */
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_oem_case_idif_any_348146']"))
						.sendKeys("N/A");
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_time_track_mandate_348146']"))
						.click();

				Thread.sleep(5000);
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_submit']")).click();

				Thread.sleep(5000);

				System.out.println("done filling form");
				formfilled = false;
			} catch (Exception e) {
				System.out.println("Exception while filing form reloading the page .." + e);
				driver.get(driver.getCurrentUrl());
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_email']"))
						.sendKeys("MSD.DeepakNitrate@embee.co.in");
				driver.findElement(By.xpath("//html")).click();
				/*
				 * driver.findElement(By.xpath("//*[@id='select2-chosen-9']")).click();
				 * //System.out.println(SubjectList.size());
				 * 
				 * driver.findElement(By.xpath("//*[@id='helpdesk_ticket_subject']")).sendKeys(
				 * Subject); Thread.sleep(5000);
				 * //System.out.println(driver.findElement(By.xpath(
				 * "//*[@id='select2-result-label-37']/span")).getText());
				 * driver.findElement(By.xpath("//*[@id='s2id_autogen9_search']")).
				 * sendKeys("Embee-MSD Managed"); System.out.println("message passed ...");
				 * 
				 * element.sendKeys(Keys.TAB); System.out.println("tab pressed");
				 * 
				 * driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				 */
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_ticket_mode_348146']")).click();
				driver.findElement(By.xpath("//*[@id='select2-results-10']/li[2]")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_nsd_member_name_348146']"))
						.click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen13_search']")).sendKeys("pou");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_source']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen28_search']")).sendKeys("ema");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_status']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen29_search']")).sendKeys("assign ");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_group_id']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen33_search']")).sendKeys("mwp s");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_responder_id']")).click();
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen34_search']")).sendKeys("re");
				Thread.sleep(2000);

				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_on_roaster_engineer_348146']"))
						.click();
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen15_search']")).sendKeys("rohi");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_department_id']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen35_search']")).sendKeys("deepak ni");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='NewTicket']/ul/li[23]/div/div[1]/div[2]")).sendKeys(desc);

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_category']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen16_search']")).sendKeys(category);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_sub_category']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen17_search']")).sendKeys(subCategory);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_custom_field_location_348146']")).click();
				driver.findElement(By.xpath("//*[@id='s2id_autogen18_search']")).sendKeys("india-w");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(5000);

				System.out.println(Subject + "first");
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_subject']")).sendKeys(arr[0]);
				System.out.println(Subject);

				Thread.sleep(4000);
				driver.findElement(By.xpath("//*[@id='select2-chosen-9']")).click();
				Thread.sleep(3000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen9_search']")).sendKeys("Embee-MSD Managed");
				System.out.println("message passed ...");
				/*
				 * element.sendKeys(Keys.TAB); System.out.println("tab pressed");
				 */
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_oem_case_idif_any_348146']"))
						.sendKeys("N/A");
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_time_track_mandate_348146']"))
						.click();

				Thread.sleep(5000);
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_submit']")).click();

				Thread.sleep(5000);

				System.out.println("done filling form in exception");
				formfilled = false;

			}
		}

		while (ticketUpdate) {
			try {
				driver.findElement(By.xpath("//*[@id='TicketPseudoReply']/a[3]")).click();
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/ul/li[2]/div[2]/div[2]")).sendKeys("Assigned");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/div[5]/span[2]/button[1]")).click();

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='TicketPseudoReply']/a[3]")).click();
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/ul/li[2]/div[2]/div[2]")).sendKeys(remarks);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/div[5]/span[2]/button[1]")).click();

				Thread.sleep(5000);
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_status']")).click();
				Thread.sleep(8000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen13_search']")).sendKeys("reso");

				// *[@id="s2id_autogen33_search"]

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(2000);

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_ticket_type']")).click();
				Thread.sleep(2000);
				driver.findElement(By.cssSelector("#select2-results-15 > li:nth-child(2)")).click();
				Thread.sleep(2000);

				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_resolution_remarks_348146']"))
						.sendKeys(remarks);

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='TimesheetTab']")).click();
				Thread.sleep(1000);
				driver.findElement(By.xpath("//*[@id='triggerAddTime']")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//*[@id='time_entry_hhmm']")).sendKeys("00:03");

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='new_timeentry-submit']")).click();

				String ticketidarr[] = driver.getCurrentUrl().split("/");
				System.out.println("Ticket id = " + ticketidarr[ticketidarr.length - 1]);
				obj.writeExcel("/home/ashish/Downloads", "Copy of CALL REPORT FOR THE MONTH OF December -2022.xlsx",
						"Sheet2", ticketidarr[ticketidarr.length - 1], rowNumber);

				Thread.sleep(4000);

				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_submit']")).click();

				// driver.findElement(By.xpath("//*[@id='save_and_new']")).click();
				Thread.sleep(10000);
				System.out.println("done updating the ticket");
				ticketUpdate = false;
			} catch (Exception e) {
				System.out.println("Exception occured after execution" + e);
				driver.get(driver.getCurrentUrl());

				driver.findElement(By.xpath("//*[@id='TicketPseudoReply']/a[3]")).click();
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/ul/li[2]/div[2]/div[2]")).sendKeys("Assigned");
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/div[5]/span[2]/button[1]")).click();

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='TicketPseudoReply']/a[3]")).click();
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/ul/li[2]/div[2]/div[2]")).sendKeys(remarks);
				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='HelpdeskNotes']/div[5]/span[2]/button[1]")).click();
				Thread.sleep(5000);

				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_status']")).click();
				Thread.sleep(8000);
				driver.findElement(By.xpath("//*[@id='s2id_autogen13_search']")).sendKeys("reso");

				Thread.sleep(3000);
				driver.findElement(By.xpath("//*[@class='select2-result-label']")).click();

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='s2id_helpdesk_ticket_ticket_type']")).click();
				Thread.sleep(2000);
				driver.findElement(By.cssSelector("#select2-results-15 > li:nth-child(2)")).click();
				Thread.sleep(2000);

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_custom_field_resolution_remarks_348146']"))
						.sendKeys(remarks);

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='TimesheetTab']")).click();
				Thread.sleep(1000);
				driver.findElement(By.xpath("//*[@id='triggerAddTime']")).click();
				Thread.sleep(5000);
				driver.findElement(By.xpath("//*[@id='time_entry_hhmm']")).sendKeys("00:03");

				Thread.sleep(2000);
				driver.findElement(By.xpath("//*[@id='new_timeentry-submit']")).click();

				String ticketidarr[] = driver.getCurrentUrl().split("/");
				System.out.println("Ticket id = " + ticketidarr[ticketidarr.length - 1]);
				obj.writeExcel("/home/ashish/Downloads", "Copy of CALL REPORT FOR THE MONTH OF December -2022.xlsx",
						"Sheet2", ticketidarr[ticketidarr.length - 1], rowNumber);

				Thread.sleep(4000);

				driver.findElement(By.xpath("//*[@id='helpdesk_ticket_submit']")).click();

				// driver.findElement(By.xpath("//*[@id='save_and_new']")).click();
				Thread.sleep(10000);
				System.out.println("done updating the ticket in exception");
				ticketUpdate = false;
			}
		}

	}

	public List<List<String>> readExcel(String filePath, String fileName, String sheetName) throws IOException {

		// Create an object of File class to open xlsx file

		File file = new File(filePath + "/" + fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook guru99Workbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			guru99Workbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of HSSFWorkbook class

			guru99Workbook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		Sheet guru99Sheet = guru99Workbook.getSheet(sheetName);

		// Find number of rows in excel file

		int rowCount = guru99Sheet.getLastRowNum() - guru99Sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it
		System.out.println("rowCount =" + rowCount);
		// rowCount = 63;
		DataFormatter fmt = new DataFormatter();

		List<List<String>> rowdataList = new ArrayList<>();

		for (int i = 0; i < rowCount + 1; i++) {

			Row row = guru99Sheet.getRow(i);
			List<String> rowdata = new ArrayList<>();
			// Create a loop to print cell values in a row

			for (int j = 0; j < row.getLastCellNum(); j++) {

				// Print Excel data in console
				// System.out.println(row.getCell(j).getCellType().toString());
				if (row.getCell(j).getCellType().toString().equalsIgnoreCase("STRING")) {
					Cell cell = row.getCell(j);
					String valueAsSeenInExcel = fmt.formatCellValue(cell);
					rowdata.add(valueAsSeenInExcel);
					// System.out.print(valueAsSeenInExcel+"|| ");
				} else {
					Cell cell = row.getCell(j);
					String valueAsSeenInExcel = fmt.formatCellValue(cell);
					rowdata.add(valueAsSeenInExcel);
					// System.out.print(valueAsSeenInExcel+"|| ");
				}
			}
			rowdataList.add(rowdata);
			// System.out.println();
		}
		return rowdataList;
	}

	public void writeExcel(String filePath, String fileName, String sheetName, String dataToWrite, int rowNumber)
			throws IOException {

		// Create an object of File class to open xlsx file
		System.out.println("Writing data.." + dataToWrite);
		File file = new File(filePath + "/" + fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook guru99Workbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			guru99Workbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of XSSFWorkbook class

			guru99Workbook = new HSSFWorkbook(inputStream);

		}

		// Read excel sheet by sheet name

		Sheet sheet = guru99Workbook.getSheet(sheetName);

		// Get the current count of rows in excel file

		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		// Get the first row from the sheet

		Row row = sheet.getRow(rowNumber);

		int columnnumber = row.getLastCellNum();

		// Create a new row and append it at last of sheet

		// Row newRow = sheet.createRow(rowCount+1);

		// Create a loop over the cell of newly created Row
		int j = columnnumber;
		// for(int j = 0; j < row.getLastCellNum(); j++){

		// Fill data in row

		Cell cell = row.createCell(j);

		cell.setCellValue(dataToWrite);

		// }

		// Close input stream

		inputStream.close();

		// Create an object of FileOutputStream class to create write data in excel file

		FileOutputStream outputStream = new FileOutputStream(file);

		// write data in the excel file

		guru99Workbook.write(outputStream);

		// close output stream

		outputStream.close();

	}

	public void autoGrab(WebDriver driver) throws Exception {
		
		try {
		driver.findElement(By.xpath("//*[@id='ui-component-check-box-input']")).click();
		driver.findElement(By.xpath("//*[@id='bulk-actions-toolbar']/div/button[4]")).click();
		Thread.sleep(20000);
		driver.get(driver.getCurrentUrl());
		autoGrab(driver);
		}catch (Exception e) {
			driver.get(driver.getCurrentUrl());
			   DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");  
			   LocalDateTime now = LocalDateTime.now();  
			   //System.out.println();  
			System.out.println("Exception "+dtf.format(now));
			Thread.sleep(20000);
			autoGrab(driver);
		}
	}

}
