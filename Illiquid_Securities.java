package BSE;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class Illiquid_Securities implements Runnable {

    public String name;
    public int minYear;
    public int maxYear;

    private Workbook resultWorkbook;
    private Sheet resultSheet;
    private int resultRowNum;

    Illiquid_Securities(String name, int minYear, int maxYear) {
        this.name = name;
        this.minYear = minYear;
        this.maxYear = maxYear;
        this.resultWorkbook = new XSSFWorkbook(); // Create the result workbook only once
        this.resultSheet = resultWorkbook.createSheet("Filtered Data");
        this.resultRowNum = 0;
    }

    public void run() {
        try {
            process();
        } catch (InterruptedException | IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void process() throws InterruptedException, IOException {
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Illiquid_Securities\\";

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));  // Reduced general wait time
        driver.get("https://www.bseindia.com/static/investors/illiquidscrips.aspx");
        driver.manage().window().maximize();

        // TreeMap to store the extracted data
        TreeMap<Integer, List<String>> map = new TreeMap<>();

        // Locate the outer table first
        WebElement outerTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.marketstartarea table")));
        // Now locate the inner table inside the outer table
        WebElement innerTable = outerTable.findElement(By.cssSelector("td table"));  // Adjusted selector to target the inner table
        // Find the tbody of the inner table
        WebElement tbody = innerTable.findElement(By.tagName("tbody"));
        // Get all rows of the inner table
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        // Variable to store the current year
        Integer currentYear = null;

        for (int j = 1; j < rows.size(); j++) {
            List<WebElement> tds = rows.get(j).findElements(By.tagName("td"));

            for (int i = 0; i < tds.size(); i++) {
                if (i == 0 && j != 0) {
                    // Extract the year from the first <td> of each row
                    currentYear = Integer.parseInt(tds.get(i).getText().trim());
                    // Initialize an empty list of URLs for this year
                    map.put(currentYear, new ArrayList<>());
                } else {
                    try {
                        // Extract the link from the <a> tag inside the <td>
                        WebElement a = tds.get(i).findElement(By.tagName("a"));
                        String url = a.getAttribute("href");

                        // Add the URL to the list corresponding to the current year
                        if (currentYear != null && map.containsKey(currentYear)) {
                            map.get(currentYear).add(url);
                        }
                    } catch (NoSuchElementException e) {
                        e.getMessage(); // Handle cases where no <a> tag is present
                    }
                }
            }
        }
        CellStyle style = resultWorkbook.createCellStyle();
        Font font = resultWorkbook.createFont();
        font.setBold(true); // Set the font bold
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER); // Center alignment

        // Create a bold font style for the header row
        CellStyle headerStyle = resultWorkbook.createCellStyle();
        Font boldFont = resultWorkbook.createFont();
        boldFont.setBold(true);
        headerStyle.setFont(boldFont);

        // Set background color for the header row to #C1F0C8
        XSSFColor headerBgColor = new XSSFColor(new byte[]{(byte) 193, (byte) 240, (byte) 200}, null);  // RGB for #C1F0C8
        ((XSSFCellStyle) headerStyle).setFillForegroundColor(headerBgColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        String headers[] = {"Sr. No.","Scrip_Code","Scrip_Name","ISIN","Effective Date"};

        Row headerRow = resultSheet.createRow(resultRowNum++);
        for( int i=0; i< headers.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        String months[] = {"Jan-Mar","Apr-Jun","Jul-Sep","Oct-Dec"};

        // Printing the TreeMap to verify
        System.out.println("\nStored URLs in TreeMap:");
        for (Map.Entry<Integer, List<String>> entry : map.entrySet()) {
            Integer currYear = entry.getKey();
            if (currYear >= minYear && currYear <= maxYear) {
                Row row = resultSheet.createRow(resultRowNum++);
                Cell yearCell = row.createCell(0);
                yearCell.setCellValue(currYear);
                resultSheet.addMergedRegion(new CellRangeAddress(resultRowNum - 1, resultRowNum - 1, 0, 4));
                yearCell.setCellStyle(style);


                System.out.println("Year : " + currYear);
                int currMonth = 0;
                for (String url : entry.getValue()) {
                    System.out.println(" - " + url);
                    if(url.contains("http")){
                        Row currMonthRow = resultSheet.createRow(resultRowNum++);
                        Cell currMonthCell = currMonthRow.createCell(0);
                        currMonthCell.setCellValue(months[currMonth]);
                        resultSheet.addMergedRegion(new CellRangeAddress(resultRowNum - 1, resultRowNum - 1, 0, 4));
                        currMonthCell.setCellStyle(style);
                        currMonth++;
                        loadPage(driver, url, wait, downloadDir);
                    }


                }
            }
        }

        // Adjust column width to fit the content and apply wrap
        for (int i = 0; i < 5; i++) {
            resultSheet.autoSizeColumn(i);  // Auto-size the columns
        }

        // Save and close the workbook after processing all files
        try (FileOutputStream fos = new FileOutputStream(downloadDir + "FilteredResults.xlsx")) {
            resultWorkbook.write(fos);
            resultWorkbook.close();
            System.out.println("All matching rows saved to FilteredResults.xlsx");
        } finally {
            driver.quit(); // Quit driver after saving the workbook
        }
    }
    void loadPage(WebDriver driver, String url, WebDriverWait wait, String downloadDir) throws InterruptedException, IOException {
        driver.get(url);
        WebElement outerTable = null;
        try{
            outerTable = driver.findElement(By.cssSelector("div.marketstartarea table"));
        }catch (Exception e ){
            e.getMessage();
        }
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");  // Timestamp for unique file names
        if (outerTable != null) {
            // Locate the outer table first
         //   WebElement outerTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.marketstartarea table")));
            // Now locate the inner table inside the outer table
            System.out.println("First layout find");


            WebElement innerTable = outerTable.findElement(By.cssSelector("td table"));  // Adjusted selector to target the inner table
            // Find the tbody of the inner table
            WebElement tbody = innerTable.findElement(By.tagName("tbody"));
            // Get all rows of the inner table
            List<WebElement> rows = tbody.findElements(By.tagName("tr"));



            for (int i = 1; i < rows.size(); i++) {
                List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
                for (int j = 1; j < tds.size(); j++) {
                    String fileName = tds.get(j).getText();
                    // Only handle Annexure.xls, Annexure.xlsx, and Annexure.zip (ignore other files)
                    if (fileName.equalsIgnoreCase("Annexure.xls") || fileName.equalsIgnoreCase("Annexure.xlsx") || fileName.equalsIgnoreCase("Annexure.zip") || fileName.equalsIgnoreCase("Annexure I.xls")) {
                        System.out.println("Found file: " + fileName);
                        WebElement a = tds.get(j).findElement(By.tagName("a"));

                        // Trigger download using Actions class to ensure proper click
                        Actions actions = new Actions(driver);
                        actions.moveToElement(a).click().perform();

                        // Append a unique timestamp to the file to avoid overwriting
                        String uniqueFileName = sdf.format(new Date()) + "_" + fileName;
                        File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);
                        if (downloadedFile != null) {
                            System.out.println("File downloaded successfully: " + downloadedFile.getName());

                            // Process zip files to extract Annexure.xls or Annexure.xlsx
                            if (fileName.equalsIgnoreCase("Annexure.zip")) {
                                String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                                if (extractedFilePath != null) {
                                    // Process the extracted Excel file (Annexure.xls or Annexure.xlsx)
                                    searchAndCopyRow(extractedFilePath, name, downloadDir + "FilteredResults.xlsx");

                                    // Clean up: Delete the zip and extracted Excel file after processing
                                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                                    Files.deleteIfExists(Paths.get(extractedFilePath));
                                }
                            } else {
                                // Process .xls or .xlsx files directly
                                searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "FilteredResults.xlsx");
                                // Clean up: Delete the Excel file after processing
                                Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                            }
                        } else {
                            System.out.println("File download failed.");
                        }
                    }
                }
            }
        }else{
            System.out.println("Second layout found.");
            // Wait for the element with id="tc52" to be visible
            WebElement downloadElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("tc52")));

            // Find the <a> tag inside the <td> and get its href attribute (the download link)
            WebElement a = downloadElement.findElement(By.tagName("a"));
          //  String fileName = a.getAttribute("href");

          //  System.out.println("Download URL: " + fileUrl);
//             Trigger download using Actions class to ensure proper click
              Actions actions = new Actions(driver);
              actions.moveToElement(a).click().perform();
              String fileName = a.getText();
            System.out.println("Found file: " + fileName);
            // Append a unique timestamp to the file to avoid overwriting
            String uniqueFileName = sdf.format(new Date()) + "_" + fileName;
            File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);
            if (downloadedFile != null) {
                System.out.println("File downloaded successfully: " + downloadedFile.getName());

                // Process zip files to extract Annexure.xls or Annexure.xlsx
                if (fileName.equalsIgnoreCase("Annexure.zip")) {
                    String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                    if (extractedFilePath != null) {
                        // Process the extracted Excel file (Annexure.xls or Annexure.xlsx)
                        searchAndCopyRow(extractedFilePath, name, downloadDir + "FilteredResults.xlsx");

                        // Clean up: Delete the zip and extracted Excel file after processing
                        Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                        Files.deleteIfExists(Paths.get(extractedFilePath));
                    }
                } else {
                    // Process .xls or .xlsx files directly
                    searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "FilteredResults.xlsx");
                    // Clean up: Delete the Excel file after processing
                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                }
            } else {
                System.out.println("File download failed.");
            }
        }
    }


    // Optimized wait for file download
    private static File waitForFileToDownload(String dirPath, String extension, int timeoutInSeconds) throws InterruptedException {
        File dir = new File(dirPath);
        File downloadedFile = null;
        int waited = 0;

        while (waited < timeoutInSeconds) {
            File[] files = dir.listFiles((d, name) -> name.toLowerCase().endsWith(extension));
            if (files != null && files.length > 0) {
                downloadedFile = files[0];
                break;
            }
            Thread.sleep(1000);  // Wait 1 second before checking again
            waited++;
        }

        if (downloadedFile != null) {
            System.out.println("Downloaded file: " + downloadedFile.getName());
        } else {
            System.out.println("No file downloaded.");
        }

        return downloadedFile;
    }

    private static String extractZipFile(String zipFilePath, String extractToDirectory) {
        try (ZipFile zipFile = new ZipFile(zipFilePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            String extractedFilePath = null;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS"); // Timestamp

            System.out.println("Listing all entries in the zip file:");

            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String fileName = entry.getName().toLowerCase().trim();

                // Log each entry in the zip file for debugging
                System.out.println("Found file in zip: " + fileName);

                // Only process .xls and .xlsx files (Annexure.xls, Annexure.xlsx, Annexures.xlsx)
                if (!entry.isDirectory() && (fileName.equals("annexure.xls") || fileName.equals("annexure.xlsx") || fileName.equals("annexures.xlsx") || fileName.equals("Annexure I.xls"))) {
                    String timestampedFileName = sdf.format(new Date()) + "_" + entry.getName();
                    extractedFilePath = extractToDirectory + File.separator + timestampedFileName;
                    File newFile = new File(extractedFilePath);

                    // Extract the Excel file from the zip
                    try (BufferedInputStream bis = new BufferedInputStream(zipFile.getInputStream(entry))) {
                        // Create parent directories if they don't exist
                        new File(newFile.getParent()).mkdirs();

                        try (FileOutputStream fos = new FileOutputStream(newFile)) {
                            byte[] bytesIn = new byte[4096];
                            int read;
                            while ((read = bis.read(bytesIn)) != -1) {
                                fos.write(bytesIn, 0, read);
                            }
                        }
                    }
                    System.out.println("Extracted Excel file: " + extractedFilePath);
                    return extractedFilePath;  // Return the path to the .xls/.xlsx file
                }
            }

            if (extractedFilePath == null) {
                System.out.println("No Annexure.xls, Annexure.xlsx, or Annexures.xlsx file found in the zip archive.");
            }
            return extractedFilePath;  // Returns the path to the extracted file or null if not found

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private void searchAndCopyRow(String excelFilePath, String keyword, String outputFilePath) {
        // Convert the keyword to lowercase to make the search case-insensitive
        String keywordLowerCase = keyword.toLowerCase();

        // Try to open the workbook (both .xls and .xlsx formats are supported)
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Iterate over all sheets in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);

                // Copy the matching rows from each sheet
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue;  // Skip the header row
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            // Compare cell value with keyword, both converted to lowercase
                            String cellValueLowerCase = cell.getStringCellValue().toLowerCase();
                            if (cellValueLowerCase.contains(keywordLowerCase)) {
                                Row resultRow = resultSheet.createRow(resultRowNum++);  // Create a new row after previous results
                                copyRow(row, resultRow);
                                break;
                            }
                        }
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Copy one row from source to destination and remove leading empty cells
    private void copyRow(Row sourceRow, Row destinationRow) {
        int destCellIndex = 0; // Destination cell index will start at 0
        boolean foundNonEmptyCell = false; // Flag to track if a non-empty cell is found

        // Iterate through each cell in the source row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);

            // Skip leading empty cells
            if (sourceCell != null && sourceCell.getCellType() != CellType.BLANK) {
                foundNonEmptyCell = true;
            }

            // Once a non-empty cell is found, start copying the cells to the destination row
            if (foundNonEmptyCell) {
                Cell destinationCell = destinationRow.createCell(destCellIndex++); // Create the cell at destination

                if (sourceCell != null) {
                    switch (sourceCell.getCellType()) {
                        case STRING:
                            destinationCell.setCellValue(sourceCell.getStringCellValue());
                            break;
                        case NUMERIC:
                            destinationCell.setCellValue(sourceCell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                            break;
                        default:
                            destinationCell.setCellValue(sourceCell.toString());
                    }
                }
            }
        }
    }



}

class Main2 {
    public static void main(String[] args) {
        Illiquid_Securities obj = new Illiquid_Securities("Rish", 2012, 2024);
        Thread thread = new Thread(obj);
        thread.start();
    }
}
