package Disciplinary_Committee_DISC;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.adobe.pdfservices.operation.PDFServices;
import com.adobe.pdfservices.operation.PDFServicesMediaType;
import com.adobe.pdfservices.operation.PDFServicesResponse;
import com.adobe.pdfservices.operation.auth.Credentials;
import com.adobe.pdfservices.operation.auth.ServicePrincipalCredentials;
import com.adobe.pdfservices.operation.io.Asset;
import com.adobe.pdfservices.operation.io.StreamAsset;
import com.adobe.pdfservices.operation.pdfjobs.jobs.ExportPDFJob;
import com.adobe.pdfservices.operation.pdfjobs.params.exportpdf.ExportPDFParams;
import com.adobe.pdfservices.operation.pdfjobs.params.exportpdf.ExportPDFTargetFormat;
import com.adobe.pdfservices.operation.pdfjobs.result.ExportPDFResult;
import com.google.api.client.util.IOUtils;

public class DISC {
    static int finalCount = 0;
    private static final Logger LOGGER = LoggerFactory.getLogger(DISC.class);

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
            String downloadDirectory = "C:\\Users\\Faizan\\eclipse-workspace\\Automation\\src\\test\\java\\Disciplinary_Committee_DISC\\";
            String resultExcelFilePath = downloadDirectory + "DISC_results.xlsx";
            // Ensure the file is empty initially
            clearExcelFile(resultExcelFilePath);

            String search_Keyword = "Shri";

            DISC obj = new DISC();
            obj.DC_1(driver, wait, search_Keyword, downloadDirectory, resultExcelFilePath, filesToDelete);
            obj.DC_3(driver, wait, search_Keyword, downloadDirectory, resultExcelFilePath, filesToDelete);
            System.out.println("Total rows found: " + finalCount);

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

    public void DC_1(WebDriver driver, WebDriverWait wait, String search_Keyword, String downloadDirectory, String resultExcelFilePath, List<String> filesToDelete) {
        String pageURL = "https://disc.icai.org/cause-list-bench-1-2023-2024/";
        String heading = "Cause List Disciplinary Committee (Bench 1)";
        List<String> all_urlList = getAllTableURL(driver, pageURL, wait);
        List<String> all_dates = getAllTableDates(driver, pageURL, wait);
        System.out.println("Total URLs: " + all_urlList.size());

        // Process all PDFs under DC_1
        for (int i = 0; i < all_urlList.size(); i++) {
            // Only pass true for the first file to add heading once for this section
            boolean isFirstPdf = (i == 0);
            processPDFSearch(all_urlList.get(i), downloadDirectory, search_Keyword, resultExcelFilePath, all_dates.get(i), heading, filesToDelete, isFirstPdf);
        }
    }

    public void DC_3(WebDriver driver, WebDriverWait wait, String search_Keyword, String downloadDirectory, String resultExcelFilePath, List<String> filesToDelete) {
        String pageURL = "https://disc.icai.org/causelist-25-b3/";
        String heading = "Cause List Disciplinary Committee (Bench 3)";
        List<String> all_urlList = getAllTableURL(driver, pageURL, wait);
        List<String> all_dates = getAllTableDates(driver, pageURL, wait);
        System.out.println("Total URLs: " + all_urlList.size());

        // Process all PDFs under DC_3
        for (int i = 0; i < all_urlList.size(); i++) {
            // Only pass true for the first file to add heading once for this section
            boolean isFirstPdf = (i == 0);
            processPDFSearch(all_urlList.get(i), downloadDirectory, search_Keyword, resultExcelFilePath, all_dates.get(i), heading, filesToDelete, isFirstPdf);
        }
    }

    // Delete file method
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

    // Process PDF using Adobe API to convert to Excel
    public static void processPDFSearch(String pdfUrl, String downloadDirectory, String searchWord, String resultExcelFilePath, String date, String heading, List<String> filesToDelete, boolean isFirstPdf) {
        try {
            // Define file paths
            String pdfFilePath = downloadDirectory + extractFileName(pdfUrl);
            String excelFilePath = downloadDirectory + extractFileNameWithoutExtension(pdfUrl) + "_adobe.xlsx";

            // Step 1: Download the PDF
            downloadPDF(pdfUrl, pdfFilePath);
            filesToDelete.add(pdfFilePath);
            filesToDelete.add(excelFilePath); // Add PDF to deletion list

            // Step 2: Convert PDF to Excel using Adobe API
            convertPDFToExcelUsingAdobeAPI(pdfFilePath, excelFilePath);
            filesToDelete.add(excelFilePath); // Add the Excel to deletion list

            // Step 3: Search and extract matching rows
            List<List<String>> matchingRows = searchInExcel(excelFilePath, searchWord);

            // Step 4: Write results to the main Excel file, passing the date
            appendResultsToExcel(matchingRows, resultExcelFilePath, new ArrayList<>(), date, heading, isFirstPdf);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    // Method to download PDF
    public static void downloadPDF(String pdfUrl, String pdfFilePath) throws Exception {
        Path pdfPath = Paths.get(pdfFilePath);
        if (Files.exists(pdfPath)) {
            Files.delete(pdfPath);
            System.out.println("Existing PDF file d.");
        }
        try (InputStream in = new URL(pdfUrl).openStream()) {
            Files.copy(in, pdfPath);
            System.out.println("PDF downloaded successfully.");
        }
    }
 // Convert PDF to Excel using Adobe PDF Services API
    public static void convertPDFToExcelUsingAdobeAPI(String pdfFilePath, String excelFilePath) throws Exception {
        Credentials credentials = new ServicePrincipalCredentials("5e9dc36b1ad641ec93e28886639405e4", "p8e-Nowt1IATwFszgEAaC4GK8p4FpRxT10xr");

        PDFServices pdfServices = new PDFServices(credentials);

        try (InputStream inputStream = Files.newInputStream(Paths.get(pdfFilePath))) {
            Asset asset = pdfServices.upload(inputStream, PDFServicesMediaType.PDF.getMediaType());

            ExportPDFParams exportPDFParams = ExportPDFParams.exportPDFParamsBuilder(ExportPDFTargetFormat.XLSX).build();
            ExportPDFJob exportPDFJob = new ExportPDFJob(asset, exportPDFParams);

            String location = pdfServices.submit(exportPDFJob);
            PDFServicesResponse<ExportPDFResult> pdfServicesResponse = pdfServices.getJobResult(location, ExportPDFResult.class);

            Asset resultAsset = pdfServicesResponse.getResult().getAsset();
            StreamAsset streamAsset = pdfServices.getContent(resultAsset);

            try (OutputStream outputStream = Files.newOutputStream(Paths.get(excelFilePath))) {
                IOUtils.copy(streamAsset.getInputStream(), outputStream);
            }
            // Log to verify the file has been created
            File generatedExcelFile = new File(excelFilePath);
            if (generatedExcelFile.exists()) {
                System.out.println("PDF converted to Excel successfully and saved at: " + excelFilePath);
            } else {
                System.out.println("Failed to create the Excel file.");
            }
        }
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
                    String cellValue = cell.toString().trim();
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
        }

        System.out.println("Total matches found: " + matchingRows.size());
        finalCount += matchingRows.size();

        // Log if no data is found
        if (matchingRows.isEmpty()) {
            System.out.println("No matching data found in the converted Excel file.");
        }
        return matchingRows;
    }


    // Append results to Excel
    // In the method to append results to Excel, modify how you handle rowData
 // Append results to Excel
 // Append results to Excel (thread-safe)
    public static synchronized void appendResultsToExcel(List<List<String>> matchingRows, String resultExcelFilePath, List<String> header, String date, String heading, boolean shouldAddHeading) throws Exception {
        if (matchingRows.isEmpty()) {
            System.out.println("No matching results to write.");
            return;
        }

        Workbook workbook;
        Sheet sheet;

        File excelFile = new File(resultExcelFilePath);

        System.out.println("Attempting to save file at: " + excelFile.getAbsolutePath());

        Path outputPath = Paths.get(resultExcelFilePath).getParent();
        if (!Files.exists(outputPath)) {
            Files.createDirectories(outputPath);
            System.out.println("Created directory: " + outputPath.toString());
        }

        // Load existing workbook or create a new one if it doesn't exist
        if (excelFile.exists()) {
            try (FileInputStream fis = new FileInputStream(excelFile)) {
                workbook = new XSSFWorkbook(fis);
                System.out.println("Opened existing Excel file: " + resultExcelFilePath);
            }
            // Use the existing sheet if it exists
            sheet = workbook.getSheet("Search Results");
            if (sheet == null) {
                sheet = workbook.createSheet("Search Results");
            }
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Search Results");
            System.out.println("Created new Excel file: " + resultExcelFilePath);
        }

        // Determine the next row number to write data
        int rowNum = sheet.getLastRowNum() + 1;
        System.out.println("Starting to append data at row: " + rowNum);

        // Add the heading and headers only if shouldAddHeading is true
        if (shouldAddHeading) {
            Row headingRow = sheet.createRow(rowNum++);
            Cell headingCell = headingRow.createCell(0);
            headingCell.setCellValue(heading);

            // Ensure at least two cells are merged for the heading
            int lastCol = Math.max(1, header.size()); // Ensure at least 1 column is available for merging
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(rowNum - 1, rowNum - 1, 0, lastCol));

            CellStyle headingStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 14);
            headingStyle.setFont(font);
            headingStyle.setAlignment(HorizontalAlignment.CENTER);
            headingStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headingCell.setCellStyle(headingStyle);
            headingRow.setHeightInPoints(25);
            System.out.println("Heading added: " + heading);

            // Add the fixed header: "S.No.", "Case No.", "Particulars of the case", "Date"
            Row headerRow = sheet.createRow(rowNum++);
            String[] fixedHeaders = {"S.No.", "Case No.", "Particulars of the case", "Date"};
            
            // Apply bold styling to the header row
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);  // Set header font to bold
            headerStyle.setFont(headerFont);

            for (int i = 0; i < fixedHeaders.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(fixedHeaders[i]);
                cell.setCellStyle(headerStyle);  // Apply the bold style to the header cells
            }
            System.out.println("Header added and set to bold.");
        }

        // Append matching rows
        for (List<String> rowData : matchingRows) {
            // Clean the rowData: remove empty cells between values
            List<String> cleanedRowData = new ArrayList<>();
            for (String cell : rowData) {
                if (!cell.isEmpty() && !cell.equals(" ")) { // Exclude empty or whitespace cells
                    cleanedRowData.add(cell);
                }
            }

            // Ensure alignment: Check if "Sl.No." column is missing, and realign the row
            if (!isNumeric(cleanedRowData.get(0))) {
                cleanedRowData.add(0, ""); // Insert a placeholder at the start if the first value is not numeric
            }

            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < cleanedRowData.size(); i++) {
                row.createCell(i).setCellValue(cleanedRowData.get(i));
            }

            // Add the date to the last column
            row.createCell(cleanedRowData.size()).setCellValue(date);
        }

        // Auto-size the columns for readability
        for (int i = 0; i < Math.max(4, matchingRows.get(0).size()) + 1; i++) {
            sheet.autoSizeColumn(i);  // Auto-size each column
        }

        // Save the workbook to the specified file
        try (FileOutputStream fos = new FileOutputStream(resultExcelFilePath)) {
            workbook.write(fos);
            System.out.println("Results written to Excel file successfully: " + resultExcelFilePath);
        } catch (Exception e) {
            System.out.println("Failed to save the Excel file: " + e.getMessage());
            e.printStackTrace();
        }
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
            // Find all <a> tags inside the 4th column
            List<WebElement> downloadLinks = row.findElements(By.cssSelector("td:nth-child(4) a"));

            // Check if the list of <a> tags is not empty
            if (!downloadLinks.isEmpty()) {
                // Get the href attribute (URL) from the first <a> tag
                String url = downloadLinks.get(0).getAttribute("href");

                // Add the URL to the ArrayList
                tempList.add(url);
            } else {
                // No <a> tag found, so log it and continue with the next row
                System.out.println("No download link found for row: " + row.getText());
            }
        }

        // Return the list of URLs
        return tempList;
    }


    
    public static List<String> getAllTableDates(WebDriver driver, String pageURL, WebDriverWait wait) {
        // Create a list to store the dates
        List<String> dateList = new ArrayList<>();

        // Open the page
        driver.get(pageURL);

        // Wait until the table is visible
        WebElement table = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("table#example")));

        // Locate all rows in the tbody of the table
        List<WebElement> rows = table.findElements(By.cssSelector("tbody tr"));

        // Loop through each row
        for (WebElement row : rows) {
            // Try to find the date in the 3rd column (assumed to be the correct column for dates)
            List<WebElement> dateElements = row.findElements(By.cssSelector("td:nth-child(3)"));

            // If a date is found, add it to the dateList
            if (!dateElements.isEmpty()) {
                String date = dateElements.get(0).getText();
                dateList.add(date);
            } else {
                // Handle case where the date is missing
                System.out.println("No date found for row: " + row.getText());
            }
        }

        // Return the list of dates
        return dateList;
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

}


