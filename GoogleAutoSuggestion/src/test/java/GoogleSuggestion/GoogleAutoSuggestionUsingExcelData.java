package GoogleSuggestion;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class GoogleAutoSuggestionUsingExcelData {

	public static void main(String[] args) throws InterruptedException, IOException {

		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		FileInputStream file = new FileInputStream(System.getProperty("user.dir") + "\\ExcelFile\\4BeatsQ1.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		driver.get("https://www.google.com/");
		driver.manage().window().maximize();

		int numberOfSheets = workbook.getNumberOfSheets();

		for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			int rowCount = sheet.getPhysicalNumberOfRows();

			for (int i = 1; i < rowCount; i++) {
				XSSFRow row = sheet.getRow(i);
				if (row != null) {
					XSSFCell cell = row.getCell(2);
					if (cell != null && cell.getCellType() == CellType.STRING) {
						String searchTerm = cell.getStringCellValue().trim();

						if (!searchTerm.isEmpty()) {
							WebElement searchBox = driver.findElement(By.name("q"));
							searchBox.clear();
							searchBox.sendKeys(searchTerm);

							Thread.sleep(3000);

							List<WebElement> list = driver
									.findElements(By.xpath("//ul[@role='listbox']//li//div[@role='option']"));
							List<String> suggestionsList1 = new ArrayList<>();

							for (WebElement suggestion : list) {
								suggestionsList1.add(suggestion.getText());
							}

							if (!suggestionsList1.isEmpty()) {
								
								String longestSuggestion = suggestionsList1.get(0);
								String shortestSuggestion = suggestionsList1.get(0);

								for (String suggestion : suggestionsList1) {
									if (suggestion.length() > longestSuggestion.length()) {
										longestSuggestion = suggestion;
									}
									if (suggestion.length() < shortestSuggestion.length()) {
										shortestSuggestion = suggestion;
									}
								}

								//Write longest and shortest suggestions to the Excel file
								XSSFCell longestCell = row.createCell(3); 
								XSSFCell shortestCell = row.createCell(4);
								longestCell.setCellValue(longestSuggestion);
								shortestCell.setCellValue(shortestSuggestion);
							}
						}
					}
				}
			}
		}

		try (FileOutputStream outFile = new FileOutputStream(
				System.getProperty("user.dir") + "\\ExcelFile\\4BeatsQ1.xlsx")) {
			workbook.write(outFile);
		}
		
		workbook.close();
        file.close();
        driver.quit();
	}

}
