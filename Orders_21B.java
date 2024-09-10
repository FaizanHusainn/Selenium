package Disciplinary_Committee_DISC;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Orders_21B implements Runnable {
	
	private String name;
	
	 Orders_21B(String name) {
		// TODO Auto-generated constructor stub
	 name = this.name;
	}

	public void run() {
		
			try {
				process();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
		
	}
	
	public List<String> getAllURL(WebDriver driver, WebDriverWait wait) {
	    List<String> urlList = new ArrayList<>();
	    
	    try {
	        // Load the webpage
	        driver.get("https://disc.icai.org/disciplinary-committee-orders/");
	        
	        // Wait until the UL containing the links is visible (you can modify the selector if needed)
	        WebElement ulElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.tb_text_wrap ul")));
	        
	        // Find all <li> elements inside the <ul>
	        List<WebElement> liElements = ulElement.findElements(By.tagName("li"));
	        
	        // Loop through each <li> and get the <a> tag inside it
	        for (WebElement li : liElements) {
	            WebElement linkElement = li.findElement(By.tagName("a")); // Find the <a> tag
	            String url = linkElement.getAttribute("href");  // Get the href attribute (URL)
	            
	            // Add the URL to the urlList
	            urlList.add(url);
	        }
	        
	        System.out.println("Extracted URLs: " + urlList);

	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	    
	    return urlList;
	}

	
	public void process() throws Exception {
		
	    WebDriver driver = new ChromeDriver();
	    driver.manage().window().maximize();
	    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
	    String downloadDirectory = "C:\\Users\\Faizan\\eclipse-workspace\\Automation\\src\\test\\java\\Disciplinary_Committee_DISC\\";
	    String resultExcelFilePath = downloadDirectory + "Order_21B_results.xlsx";
	    clearExcelFile(resultExcelFilePath);


	    List<String> allURLs = getAllURL(driver, wait); // Assuming this method gets all URLs
	    System.out.println("Total URLs: " + allURLs.size());
	    String name = "Price Waterhouse"; // Name to match in the case

	    // Create an Excel workbook and sheet
	    Workbook workbook = new XSSFWorkbook();
	    Sheet sheet = workbook.createSheet("DISC Results");

	    // Define the row number to write data in Excel
	    int rowNum = 0;
	  //*[@id="themify_builder_content-6034"]/div[2]/div/div/div/div/div/h4
	    for (String current_url : allURLs) {
	        driver.get(current_url);

	        // Find the heading (H3 or H4)
	        WebElement heading = null;
	        if (!driver.findElements(By.tagName("h3")).isEmpty()) {
	            heading = driver.findElement(By.tagName("h3"));
	        } else if (!driver.findElements(By.tagName("h4")).isEmpty()) {
	            heading = driver.findElement(By.tagName("h4"));
	        }
	        String headingText = heading != null ? heading.getText() : "No Heading";

	        // Create a heading row only once per URL
	        Row headingRow = sheet.createRow(rowNum++);
	        Cell headingCell = headingRow.createCell(0);
	        headingCell.setCellValue(headingText);
	        // Apply font size and bold to heading
	        CellStyle headingStyle = workbook.createCellStyle();
	        Font font = workbook.createFont();
	        font.setBold(true);
	        font.setFontHeightInPoints((short) 14);
	        headingStyle.setFont(font);
	     
	        headingStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	        headingCell.setCellStyle(headingStyle);
	        headingRow.setHeightInPoints(25);

	        // Merge heading across 4 columns (A-D)
	        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(rowNum - 1, rowNum - 1, 0, 3));

	        // Create a header row once per URL
	        Row headerRow = sheet.createRow(rowNum++);
	        String[] headers = {"S.No.", "Case No.", "Particulars of the case", "Date of Order"};
	        CellStyle headerStyle = workbook.createCellStyle();
	        Font headerFont = workbook.createFont();
	        headerFont.setBold(true); // Make header bold
	        headerStyle.setFont(headerFont);

	        for (int i = 0; i < headers.length; i++) {
	            Cell headerCell = headerRow.createCell(i);
	            headerCell.setCellValue(headers[i]);
	            headerCell.setCellStyle(headerStyle);
	        }

	        // Find the table rows
	        List<WebElement> rows = driver.findElements(By.xpath("//table[@id='example']/tbody/tr"));

	     // Inside your for loop where you iterate over each row
	        for (WebElement row : rows) {
	            // Get the "Particulars of the case" column text
	            String particularsOfCase = row.findElement(By.xpath("td[3]")).getText();
	            
	            // If the name matches, store the data in Excel
	            if (particularsOfCase.toLowerCase().contains(name.toLowerCase())) {
	                // Create a new row in the Excel sheet
	                Row excelRow = sheet.createRow(rowNum++);

	                // Write the table row data into Excel
	                excelRow.createCell(0).setCellValue(row.findElement(By.xpath("td[1]")).getText()); // Sl.No.

	                // Case No. with multiple links
	                Cell caseNoCell = excelRow.createCell(1);

	                // First, get all the <a> tags inside td[2] directly
	                List<WebElement> caseNoLinkElements = row.findElements(By.xpath("td[2]/a"));

	                // If no <a> tags are found directly, check inside <p> tags within td[2]
	                if (caseNoLinkElements.isEmpty()) {
	                    caseNoLinkElements = row.findElements(By.xpath("td[2]/p/a"));
	                }

	                if (!caseNoLinkElements.isEmpty()) {
	                    // To store concatenated case number text with URLs
	                    StringBuilder caseNoWithLinks = new StringBuilder();

	                    for (WebElement linkElement : caseNoLinkElements) {
	                        String caseNoText = linkElement.getText();
	                        String caseNoLink = linkElement.getAttribute("href");

	                        // Append case number text with the link embedded
	                        caseNoWithLinks.append(caseNoText).append(", ");

	                        // Add hyperlink to the case number
	                        CreationHelper creationHelper = workbook.getCreationHelper();
	                        org.apache.poi.ss.usermodel.Hyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.URL);
	                        hyperlink.setAddress(caseNoLink);

	                        // Create a new cell for each case number with the hyperlink
	                        Cell newCaseNoCell = excelRow.createCell(1);
	                        newCaseNoCell.setCellValue(caseNoText);
	                        newCaseNoCell.setHyperlink(hyperlink);

	                        // Style the link (optional)
	                        CellStyle linkStyle = workbook.createCellStyle();
	                        Font linkFont = workbook.createFont();
	                        linkFont.setUnderline(Font.U_SINGLE);
	                        linkFont.setColor(IndexedColors.BLUE.getIndex());
	                        linkStyle.setFont(linkFont);
	                        newCaseNoCell.setCellStyle(linkStyle);
	                    }

	                    // Remove the last comma and space from the concatenated text
	                    if (caseNoWithLinks.length() > 0) {
	                        caseNoWithLinks.setLength(caseNoWithLinks.length() - 2);  // Remove trailing comma and space
	                    }

	                    // Set the value in the first case number cell as concatenated case numbers
	                    caseNoCell.setCellValue(caseNoWithLinks.toString());

	                } else {
	                    // If no <a> tags are found, just set the text
	                    caseNoCell.setCellValue(row.findElement(By.xpath("td[2]")).getText());
	                }

	                // Write other data
	                excelRow.createCell(2).setCellValue(particularsOfCase); // Particulars of the case
	                excelRow.createCell(3).setCellValue(row.findElement(By.xpath("td[4]")).getText()); // Date of Order
	            }
	        }



	        // Adjust column width and row height for readability
	        for (int i = 0; i < 4; i++) {
	            sheet.autoSizeColumn(i);
	        }
	        break;
	    }

	    // Save the Excel file
	    try (FileOutputStream fileOut = new FileOutputStream(downloadDirectory + "Order_21B_results.xlsx")) {
	        workbook.write(fileOut);
	    }

	    // Close the workbook and driver
	    workbook.close();
	    driver.quit();
	}

	
	public static void clearExcelFile(String resultExcelFilePath) throws Exception {
        File excelFile = new File(resultExcelFilePath);
        if (excelFile.exists()) {
            if (excelFile.delete()) {
                System.out.println("Existing result Excel file deleted.");
            } else {
                throw new Exception("Failed to delete existing result Excel file.");
            }
        }
    }
	
}

class Main_cl{
	public static void main(String args[]) {
		Orders_21B obj = new Orders_21B("abc");
		Thread thread_obj = new Thread(obj);
		thread_obj.start();
		
	}
	
}


