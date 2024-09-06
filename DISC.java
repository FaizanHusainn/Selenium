package Disciplinary_Committee_DISC;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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

import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.PageIterator;
import technology.tabula.RectangularTextContainer;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;


public class DISC {
	static int finalCount = 0;

	 public static void main(String[] args) {

	        WebDriver driver = null;
	        try {
	            // Initialize ChromeDriver
	            driver = new ChromeDriver();
	            driver.manage().window().maximize();
	            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
	            
	         // List to track downloaded PDF and Excel files for later deletion
	            List<String> filesToDelete = new ArrayList<>();

	            // Parameters
	            String downloadDirectory = "C:\\Users\\Faizan\\eclipse-workspace\\Automation\\src\\test\\java\\Disciplinary_Committee_DISC\\"; // Base directory
	            String resultExcelFilePath = downloadDirectory + "DISC_results.xlsx";
	            // Ensure the file is empty initially
	            clearExcelFile(resultExcelFilePath);
	            
	            String search_Keyword = "Shri";
	            
	            DISC obj = new DISC();
	           // obj.DC_1(driver, wait, search_Keyword, downloadDirectory, resultExcelFilePath, filesToDelete);
	            obj.DC_2(driver, wait, search_Keyword, downloadDirectory, resultExcelFilePath, filesToDelete);
	            System.out.println("TotalDigits Rows found: "+ finalCount);
	            
	            for (String filePath : filesToDelete) {
	                deleteFile(filePath);
	            }

	        } catch (Exception e) {
	            e.printStackTrace();
	        } finally {
	            System.out.println("Execution complete");
	            if (driver != null) {
	                driver.quit();
	            }
	        }
	    }
	 
	 public void DC_1(WebDriver driver,WebDriverWait wait, String search_Keyword,String downloadDirectory, String resultExcelFilePath, List<String> filesToDelete) {
			
		 String pageURL = "https://disc.icai.org/cause-list-bench-1-2023-2024/";
		
		 String heading =  "Cause List Disciplinary Committee (Bench 1)";
		 List<String> all_urlList = getAllTableURL(driver, pageURL,wait);
		 List<String> all_dates = getAllTableDates(driver, pageURL,wait);
		 System.out.println("TotalDigits URlsL" + all_urlList.size());
		 for(int i =0;  i< all_urlList.size(); i++) {
			processPDFSearch(all_urlList.get(i),downloadDirectory, search_Keyword, resultExcelFilePath, all_dates.get(i),heading, filesToDelete);
		 }
		 		 
	 }
	 public void DC_2(WebDriver driver,WebDriverWait wait, String search_Keyword,String downloadDirectory, String resultExcelFilePath, List<String> filesToDelete) {
			
		 String pageURL = "https://disc.icai.org/causelist-25-b3/";
		
		 String heading =  "Cause List Disciplinary Committee (Bench 2)";
		 List<String> all_urlList = getAllTableURL(driver, pageURL,wait);
		 List<String> all_dates = getAllTableDates(driver, pageURL,wait);
		 System.out.println("TotalDigits URlsL" + all_urlList.size());
		 for(int i =0;  i< all_urlList.size(); i++) {
			processPDFSearch(all_urlList.get(i),downloadDirectory, search_Keyword, resultExcelFilePath, all_dates.get(i),heading, filesToDelete);
		 }
		 		 
	 }
	 
	 public static void deleteFile(String filePath) {
		    try {
		        Path path = Paths.get(filePath);
		        if (Files.exists(path)) {
		            Files.delete(path);
		            System.out.println("File deleted successfully: " + filePath);
		        } else {
		            System.out.println("File not found: " + filePath);
		        }
		    } catch (Exception e) {
		        System.out.println("Failed to delete file: " + filePath);
		        e.printStackTrace();
		    }
		}
	 
	 public static void processPDFSearch(String pdfUrl, String downloadDirectory, String searchWord, String resultExcelFilePath, String date, String heading, List<String> filesToDelete) {
		    try {
		        // Define file paths
		        String pdfFilePath = downloadDirectory + extractFileName(pdfUrl);
		        String excelFilePath = downloadDirectory + extractFileNameWithoutExtension(pdfUrl) + "_tabula.xlsx";

		        // Step 1: Download the PDF
		        downloadPDF(pdfUrl, pdfFilePath);
		        filesToDelete.add(pdfFilePath);  // Add PDF to deletion list

		        // Step 2: Convert PDF to Excel using Tabula
		        List<String> header = convertPDFToExcelUsingTabula(pdfFilePath, excelFilePath);

		        // Step 3: Search and extract matching rows
		        List<List<String>> matchingRows = searchInExcel(excelFilePath, searchWord);

		        // Step 4: Write results to the main Excel file, passing the date
		        appendResultsToExcel(matchingRows, resultExcelFilePath, header, date, heading);

		    } catch (Exception e) {
		        e.printStackTrace();
		    }
		}

	    // Method to download PDF
	    public static void downloadPDF(String pdfUrl, String pdfFilePath) throws Exception {
	        Path pdfPath = Paths.get(pdfFilePath);
	        if (Files.exists(pdfPath)) {
	            Files.delete(pdfPath);
	            System.out.println("Existing PDF file deleted.");
	        }
	        try (InputStream in = new URL(pdfUrl).openStream()) {
	            Files.copy(in, pdfPath);
	            System.out.println("PDF downloaded successfully.");
	        }
	    }

	    // Method to convert PDF to Excel using Tabula
	    public static List<String> convertPDFToExcelUsingTabula(String pdfFilePath, String excelFilePath) throws Exception {
	        List<String> header = new ArrayList<>();

	        try (PDDocument document = PDDocument.load(new File(pdfFilePath));
	             Workbook workbook = new XSSFWorkbook()) {

	            ObjectExtractor extractor = new ObjectExtractor(document);
	            SpreadsheetExtractionAlgorithm algorithm = new SpreadsheetExtractionAlgorithm();
	            Sheet sheet = workbook.createSheet("Extracted Data");

	            int rowNum = 0;
	            PageIterator pages = extractor.extract(); // Correct usage of PageIterator
	            while (pages.hasNext()) {
	                Page page = pages.next();
	                List<Table> tables = algorithm.extract(page);

	                for (Table table : tables) {
	                    for (List<RectangularTextContainer> row : table.getRows()) {
	                        Row excelRow = sheet.createRow(rowNum++);
	                        for (int i = 0; i < row.size(); i++) {
	                            String cellValue = row.get(i).getText();
	                            if (rowNum == 1) {
	                                header.add(cellValue); // Capture the header from the first row
	                            }
	                            excelRow.createCell(i).setCellValue(cellValue);
	                        }
	                    }
	                }
	            }

	            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
	                workbook.write(fos);
	            }

	            System.out.println("Tables extracted and saved to Excel successfully.");
	        }

	        return header;
	    }
	    
	    // Utility method to normalize text
	    public static String normalizeText(String text) {
	        if (text == null) return "";
	        return text.trim().replaceAll("\\s+", " ");
	    }
	    // Utility method to extract file name from URL
	    public static String extractFileName(String url) {
	        return url.substring(url.lastIndexOf('/') + 1);
	    }

	    // Utility method to extract file name without extension from URL
	    public static String extractFileNameWithoutExtension(String url) {
	        String fileName = extractFileName(url);
	        return fileName.contains(".") ? fileName.substring(0, fileName.lastIndexOf('.')) : fileName;
	    }

	    // Method to search in Excel
	    public static List<List<String>> searchInExcel(String excelFilePath, String searchWord) throws Exception {
	        List<List<String>> matchingRows = new ArrayList<>();

	        try (FileInputStream fis = new FileInputStream(excelFilePath);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            Sheet sheet = workbook.getSheetAt(0);
	            Pattern pattern = Pattern.compile(Pattern.quote(searchWord), Pattern.CASE_INSENSITIVE);

	            int rowCount = 0;
	            for (Row row : sheet) {
	                rowCount++;
	                List<String> rowData = new ArrayList<>();
	                boolean isMatch = false;

	                for (Cell cell : row) {
	                    String cellValue = normalizeText(cell.toString());
	                    rowData.add(cellValue);

	                    Matcher matcher = pattern.matcher(cellValue);
	                    if (matcher.find()) {
	                        isMatch = true;
	                    }
	                }

	                if (isMatch && rowCount > 1) { // Skip the header row
	                    matchingRows.add(rowData);
	                    System.out.println("Match found in row " + rowCount + ": " + rowData);
	                }
	            }
	            System.out.println("Total matches found: " + matchingRows.size());
	            finalCount+= matchingRows.size();
	        }
	        return matchingRows;
	    }

	    // Method to append results to Excel
	 // Modified appendResultsToExcel to center and merge the heading at the top of the Excel file
	    public static void appendResultsToExcel(List<List<String>> matchingRows, String resultExcelFilePath, List<String> header, String date, String heading) throws Exception {
	        if (matchingRows.isEmpty()) {
	            System.out.println("No matching results to write.");
	            return;
	        }

	        Workbook workbook;
	        Sheet sheet;

	        File excelFile = new File(resultExcelFilePath);

	        // If the Excel file exists, open it; otherwise create a new one
	        if (excelFile.exists()) {
	            try (FileInputStream fis = new FileInputStream(excelFile)) {
	                workbook = new XSSFWorkbook(fis);
	            }
	            sheet = workbook.getSheetAt(0);
	        } else {
	            workbook = new XSSFWorkbook();
	            sheet = workbook.createSheet("Search Results");
	        }

	        // Track already written rows for the current PDF to avoid duplicates during its processing
	        List<List<String>> writtenRowsInCurrentPDF = new ArrayList<>();

	        int rowNum = sheet.getLastRowNum() + 1;

	        // Add heading at the top (first row) if it's a new file
	        if (rowNum == 0) {
	            // Write the heading in the first row
	            Row headingRow = sheet.createRow(rowNum++);
	            Cell headingCell = headingRow.createCell(0);
	            headingCell.setCellValue(heading);

	            // Merge the heading across all columns (adjust based on the number of columns in the header)
	            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, header.size()));

	            // Create a style to center the heading
	            CellStyle headingStyle = workbook.createCellStyle();
	            Font font = workbook.createFont();
	            font.setBold(true);
	            font.setFontHeightInPoints((short) 14); // Set font size for the heading
	            headingStyle.setFont(font);
	            headingStyle.setAlignment(HorizontalAlignment.CENTER); // Center horizontally
	            headingStyle.setVerticalAlignment(VerticalAlignment.CENTER); // Center vertically

	            headingCell.setCellStyle(headingStyle);
	            headingRow.setHeightInPoints(25); // Adjust row height for better visibility
	        }

	        // Add header row (column names) after the heading, if the file is new
	        if (rowNum == 1) {
	            Row headerRow = sheet.createRow(rowNum++);
	            for (int i = 0; i < header.size(); i++) {
	                Cell cell = headerRow.createCell(i);
	                cell.setCellValue(header.get(i));
	                CellStyle style = workbook.createCellStyle();
	                Font font = workbook.createFont();
	                font.setBold(true);
	                style.setFont(font);
	                cell.setCellStyle(style);
	            }

	            // Add the "Date" column header
	            Cell dateCell = headerRow.createCell(header.size());
	            dateCell.setCellValue("Date");
	            CellStyle style = workbook.createCellStyle();
	            Font font = workbook.createFont();
	            font.setBold(true);
	            style.setFont(font);
	            dateCell.setCellStyle(style);
	        }

	        boolean dateAdded = false; // Flag to ensure the date is added only once per PDF

	        // Write each matching row, ensuring no duplicates for this specific PDF extraction
	        for (List<String> rowData : matchingRows) {
	            if (!writtenRowsInCurrentPDF.contains(rowData)) {

	                // Remove any empty cells between values
	                rowData.removeIf(String::isEmpty);

	                Row row = sheet.createRow(rowNum++);

	                // Write the row data to the Excel file
	                for (int i = 0; i < rowData.size(); i++) {
	                    row.createCell(i).setCellValue(rowData.get(i));
	                }

	                // Add the date to the first matched row, if not already added
	                if (!dateAdded) {
	                    row.createCell(rowData.size()).setCellValue(date);
	                    dateAdded = true;
	                }

	                // Add the row to the list to prevent writing it again during this PDF's processing
	                writtenRowsInCurrentPDF.add(rowData);
	            }
	        }

	        // Auto-size the columns for better readability
	        for (int i = 0; i < header.size() + 1; i++) { // Include the Date column in auto-sizing
	            sheet.autoSizeColumn(i);
	        }

	        // Save the Excel file
	        try (FileOutputStream fos = new FileOutputStream(resultExcelFilePath)) {
	            workbook.write(fos);
	        }

	        System.out.println("Results written to Excel file successfully: " + resultExcelFilePath);
	    }

	    // Helper method to check if a string is numeric
	    public static boolean isNumeric(String str) {
	        if (str == null || str.isEmpty()) {
	            return false;
	        }
	        try {
	            Double.parseDouble(str);
	            return true;
	        } catch (NumberFormatException e) {
	            return false;
	        }
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
	    public static List<String> getAllTableURL(WebDriver driver, String pageURL, WebDriverWait wait) {
		    // Create a list to store the URLs
		    List<String> tempList = new ArrayList<>();
		    
		    // Open the page
		    driver.get(pageURL);
		    
		    // Wait until the table is visible
		    WebElement table = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("table#example")));
		    
		    // Locate all rows in the tbody of the table
		    List<WebElement> rows = table.findElements(By.cssSelector("tbody tr"));
		    
		    // Loop through each row
		    for (WebElement row : rows) {
		        // Locate the <a> tag inside the "View/Download" column (4th column)
		        WebElement downloadLink = row.findElement(By.cssSelector("td:nth-child(4) a"));
		        
		        // Get the href attribute (URL) from the <a> tag
		        String url = downloadLink.getAttribute("href");
		        
		        // Add the URL to the ArrayList
		        tempList.add(url);
		    }
		    
		    // Return the list of URLs
		    return tempList;
		}
	    
	    public static List<String> getAllTableDates(WebDriver driver, String pageURL, WebDriverWait wait) {
		    // Create a list to store the URLs
		    List<String> dateList = new ArrayList<>();
		    
		    // Open the pag
		    
		    // Wait until the table is visible
		    WebElement table = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("table#example")));
		    
		    // Locate all rows in the tbody of the table
		    List<WebElement> rows = table.findElements(By.cssSelector("tbody tr"));
		    
		    // Loop through each row
		    for (WebElement row : rows) {
		        // Locate the <a> tag inside the "View/Download" column (4th column)
		        WebElement downloadLink = row.findElement(By.cssSelector("td:nth-child(3)"));
		        
		        // Get the href attribute (URL) from the <a> tag
		        String date = downloadLink.getText();
		        
		        // Add the URL to the ArrayList
		        dateList.add(date);
		    }
		    
		    // Return the list of URLs
		    return dateList;
		}

	 

}
