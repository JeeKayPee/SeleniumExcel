package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import common.ConfigurationExcel;

public class ReadWriteExcel {
	static WebDriver driver;
	static WebDriverWait wait;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static XSSFCell cell;
	static File src;

	static void ExcelUtil() throws IOException {
		String fileName = "TestData.xlsx";
		String filePath = System.getProperty("user.dir") + File.separator + "dataSource";

		// import the excel sheet
		src = new File(filePath + File.separator + fileName);
		// Load the file
		FileInputStream fileInputStream = new FileInputStream(src);

		workbook = new XSSFWorkbook(fileInputStream); // only for xlsx
		sheet = workbook.getSheet("Credentials");
	}

	static String getUserName(int i) {
		cell = sheet.getRow(i).getCell(0);
		cell.setCellType(CellType.STRING);
		return cell.getStringCellValue();
	}

	static String getPassword(int i) {
		cell = sheet.getRow(i).getCell(1);
		cell.setCellType(CellType.STRING);
		return cell.getStringCellValue();
	}

	public static void main(String[] args) throws Exception {

		driver = ConfigurationExcel.createChromeDriver();
		wait = new WebDriverWait(driver, 60);
		driver.get(ConfigurationExcel.ADMIN_URL);

		ExcelUtil();

		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			WebElement userTextBox = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("log")));
			userTextBox.sendKeys(getUserName(i));

			WebElement pwdTextBox = driver.findElement(By.name("pwd"));
			pwdTextBox.sendKeys(getPassword(i));
			pwdTextBox.submit();

			// do some test also - pass/fail - we can write excel

			String xFinder = String.format("//span[text()='%s']", getUserName(i));
			WebElement howdy = driver.findElement(By.xpath(xFinder));

			String message;
			if (howdy.isDisplayed()) {
				message = "Pass";
			} else {
				message = "Fail";
			}
			sheet.getRow(i).createCell(2).setCellValue(message);

			FileOutputStream fileOutput = new FileOutputStream(src);

			workbook.write(fileOutput);
			fileOutput.close();

			WebElement logout = driver.findElement(By.xpath("//*[text()='Log Out']"));
			driver.get(logout.getAttribute("href"));

		}

		driver.quit();

	}

}
