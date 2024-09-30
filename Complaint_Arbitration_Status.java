package BSE;

import org.apache.poi.ss.usermodel.*;
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

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class Complaint_Arbitration_Status implements Runnable {

    public String name;
//    private Workbook resultWorkbook;
//    private Sheet resultSheet;
//    private int resultRowNum;

    Complaint_Arbitration_Status(String name) {
        this.name = name;
//        this.resultWorkbook = new XSSFWorkbook(); // Create the result workbook only once
//        this.resultSheet = resultWorkbook.createSheet("Filtered Data");
        //       this.resultRowNum = 0;
    }

    @Override
    public void run() {
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.bseindia.com/investors/invgrievstats.aspx?expandable=2");
        driver.manage().window().maximize();

        // Define the map with predefined keys and empty lists as values
        Map<String, List<String>> urlMap = new HashMap<>();
        urlMap.put("complaints received from clients against trading members", new ArrayList<>());
        urlMap.put("complaints received from investors against listed companies", new ArrayList<>());
        urlMap.put("redressal of complaints lodged by clients against trading members", new ArrayList<>());
        urlMap.put("redressal of complaints lodged by investors against listed companies", new ArrayList<>());
        urlMap.put("arbitrator-wise arbitrator proceedings", new ArrayList<>());
        urlMap.put("penal actions against trading members", new ArrayList<>());
        urlMap.put("disposal of arbitration proceedings", new ArrayList<>());

        // Locate the main table
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement mainTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='divID']/table")));

        // Extracting the Years
        List<String> completeYears = new ArrayList<>();
        List<WebElement> year = mainTable.findElements(By.tagName("p"));
        for (int i = 0; i < year.size(); i++) {
            if (year.get(i).getText().contains("Data for year")) {
                completeYears.add(year.get(i).getText());
            }
        }
        System.out.println("Complete Years :" + completeYears.size());

        // Locate the tbody within the main table
        WebElement tbody = mainTable.findElement(By.tagName("tbody"));
        // Get all <tr> elements
        List<WebElement> trs = tbody.findElements(By.tagName("tr"));
        // We're targeting nested tables inside the 3rd <tr> as you mentioned
        if (trs.size() > 2) {  // Ensure there are enough <tr> elements
            WebElement targetTd = trs.get(2).findElement(By.tagName("td")); // Get the <td> from the 3rd <tr>
            // Now find all the nested tables inside this <td>
            List<WebElement> nestedTables = targetTd.findElements(By.tagName("table"));
            System.out.println("No of tables :" + nestedTables.size());

            // Iterate over each nested table
            for (int i = 1; i < nestedTables.size(); i++) {
                // Process the rows of each nested table
                List<WebElement> rows = nestedTables.get(i).findElements(By.tagName("tr"));

                for (WebElement row : rows) {
                    // Get all <td> elements in the current row
                    List<WebElement> tds = row.findElements(By.tagName("td"));

                    for (WebElement td : tds) {
                        WebElement anchor = null;

                        // Try different HTML structures to find the <a> tag
                        try {
                            // First attempt: direct <a> inside td
                            anchor = td.findElement(By.tagName("a"));
                        } catch (Exception e1) {
                            try {
                                // Second attempt: td p span font font a
                                anchor = td.findElement(By.cssSelector("p span font font a"));
                            } catch (Exception e2) {
                                try {
                                    // Third attempt: td p span a
                                    anchor = td.findElement(By.cssSelector("p span a"));
                                } catch (Exception e3) {
                                    try {
                                        // Fourth attempt: td u a
                                        anchor = td.findElement(By.cssSelector("u a"));
                                    } catch (Exception e4) {
                                        try {
                                            anchor = td.findElement(By.cssSelector("p font font span a"));
                                        } catch (Exception e5) {
                                            try {
                                                anchor = td.findElement(By.cssSelector("p font span a"));
                                            } catch (Exception e6) {
                                                try {
                                                    anchor = td.findElement(By.cssSelector("p font a"));
                                                } catch (Exception e7) {
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // If an <a> tag is found, print its text and URL
                        if (anchor != null && anchor.getText() != null) {
                            String text = anchor.getText().toLowerCase();  // Convert text to lower case
                            String url = anchor.getAttribute("href");      // Get the URL

                            // Check if the text contains any of the predefined keys and add the URL to the map
                            if (text.contains("complaints received from clients against trading members")) {
                                urlMap.get("complaints received from clients against trading members").add(url);
                            } else if (text.contains("complaints received from investors against listed companies")) {
                                urlMap.get("complaints received from investors against listed companies").add(url);
                            } else if (text.contains("redressal of complaints lodged by clients against trading members")) {
                                urlMap.get("redressal of complaints lodged by clients against trading members").add(url);
                            } else if (text.contains("redressal of complaints lodged by investors against listed companies")) {
                                urlMap.get("redressal of complaints lodged by investors against listed companies").add(url);
                            } else if (text.contains("arbitrator-wise arbitrator proceedings")) {
                                urlMap.get("arbitrator-wise arbitrator proceedings").add(url);
                            } else if (text.contains("penal actions against trading members")) {
                                urlMap.get("penal actions against trading members").add(url);
                            } else if (text.contains("disposal of arbitration proceedings")) {
                                urlMap.get("disposal of arbitration proceedings").add(url);
                            }

                            //  Print the text and the URL for debugging purposes
//                            System.out.println("Text: " + text);
                            // System.out.println("URL: " + url);
                        }
                    }
                }
            }
        } else {
            System.out.println("The main table doesn't contain enough rows to access the 3rd <tr>.");
        }

        // Create an ExecutorService with a fixed thread pool
        //    ExecutorService executor = Executors.newFixedThreadPool(1);

//        // Submit both tasks to the executor using lambda expressions
//        executor.submit(() -> {
//            try {
//                compalints_Against_TM(name, urlMap.get("complaints received from clients against trading members"));
//            } catch (InterruptedException | IOException e) {
//                e.printStackTrace();
//            }
//        });
//
//        executor.submit(() -> {
//            try {
//                compalaints_Listed_Companies(name, urlMap.get("complaints received from investors against listed companies"));
//            } catch (InterruptedException | IOException e) {
//                e.printStackTrace();
//            }
//        });
//        executor.submit(() -> {
//            try {
//                redressal_complaint_TM(name, urlMap.get("redressal of complaints lodged by clients against trading members"));
//            } catch (InterruptedException | IOException e) {
//                e.printStackTrace();
//            }
//        });
//        executor.submit(() -> {
//            try {
//                redessal_complaint_LC(name, urlMap.get("redressal of complaints lodged by investors against listed companies"));
//            } catch (InterruptedException | IOException e) {
//                e.printStackTrace();
//            }
//        });
//
//        // Shut down the executor service
//        executor.shutdown();

        try {
            redressal_complaint_TM(name, urlMap.get("redressal of complaints lodged by clients against trading members"));
        } catch (InterruptedException | IOException e) {
            e.printStackTrace();
        }
        // After processing, print the map keys and their corresponding URL lists
//        for (String key : urlMap.keySet()) {
//            System.out.println("Key: " + key + " -> URLs Count: " + urlMap.get(key).size());
//            System.out.println("URLs for Key: " + key);
//            for (String url : urlMap.get(key)) {
//                System.out.println(url);
//            }
//        }

        driver.quit();
    }

    void arbitrator_Procedings(String name, List<String> urls) throws InterruptedException, IOException{
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Complaint_Arbitration_Status\\";

        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Filter Data");
        int resultRowNum = 0;

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        ChromeDriver driver = new ChromeDriver(options);
        String[] headers = {""};
    }

    void redessal_complaint_LC (String name, List<String> urls) throws InterruptedException, IOException{
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Complaint_Arbitration_Status\\";

        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Filter Data");
        int resultRowNum = 0;

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        ChromeDriver driver = new ChromeDriver(options);

        String[] headers ={"Sr. No.","Name of the Company","If Suspended than Date of Suspension","$   No. of Folios of Shareholders","Received","Redressed through Exchange",
                "Non Actionable","Arbitration Advised","Pending for Redressal with Exchange"};
        for(int i=0; i< urls.size(); i++){
            driver.get(urls.get(i));
            String uniqueFileName = urls.get(i);
            File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);

            if (downloadedFile != null) {
                System.out.println("File downloaded successfully: " + downloadedFile.getName());

                // Process zip files to extract
                if (uniqueFileName.contains(".zip")) {
                    String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                    if (extractedFilePath != null) {
                        // Process the extracted Excel file
                        searchAndCopyRow(extractedFilePath, name, downloadDir + "Redressal_of_complaints_lodged_by_investor_against_Listed_Companies.xlsx", resultSheet,resultRowNum, resultWorkbook, headers);

                        // Clean up: Delete the zip and extracted Excel file after processing
//                        Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
//                        Files.deleteIfExists(Paths.get(extractedFilePath));
                    }
                } else {
                    // Process .xls or .xlsx files directly
                    searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "Redressal_of_complaints_lodged_by_investor_against_Listed_Companies.xlsx",resultSheet, resultRowNum, resultWorkbook, headers);
                    // Clean up: Delete the Excel file after processing

                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                }
            } else {
                System.out.println("File download failed.");
            }
        }
        driver.quit();

    }

    void redressal_complaint_TM(String name, List<String> urls) throws InterruptedException, IOException{
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Complaint_Arbitration_Status\\";

        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Filter Data");
        int resultRowNum = 0;

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        ChromeDriver driver = new ChromeDriver(options);
        String[] headers = {"Name of the TM","Clg. No.","No.of UCCs at the beginning of the year","No.of Complaints received against the TM","Resolved through the Exchange/IGRC",
                "Non actionable","Advised by IGRC to refer to Arbitra-tion, if so desired","Pending for redressal with Exchange","No.of Arbitration filed by Clients","Decided by the Arbitrators",
                "Decided by Arbitrators in favour of the clients","Pending for Redressal with Arbitrators","No. of Active clients","% of No. of complaints as against No. of active clients","% of complaints resolved as against complaints received"};

        for(int i=0; i< urls.size(); i++){
            driver.get(urls.get(i));
            String uniqueFileName = urls.get(i);
            File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);

            if (downloadedFile != null) {
                System.out.println("File downloaded successfully: " + downloadedFile.getName());

                // Process zip files to extract
                if (uniqueFileName.contains(".zip")) {
                    String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                    if (extractedFilePath != null) {
                        // Process the extracted Excel file
                        searchAndCopyRow(extractedFilePath, name, downloadDir + "Redressal_of_complaints_lodged_by_clients_against_trading_members.xlsx", resultSheet,resultRowNum, resultWorkbook, headers);

                        // Clean up: Delete the zip and extracted Excel file after processing
                        Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                        Files.deleteIfExists(Paths.get(extractedFilePath));
                    }
                } else {
                    // Process .xls or .xlsx files directly
                    searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "Redressal_of_complaints_lodged_by_clients_against_trading_members.xlsx",resultSheet, resultRowNum, resultWorkbook, headers);
                    // Clean up: Delete the Excel file after processing
                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                }
            } else {
                System.out.println("File download failed.");
            }
        }
        driver.quit();


    }

    void compalaints_Listed_Companies(String name, List<String> urls) throws InterruptedException, IOException{
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Complaint_Arbitration_Status\\";

        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Filter Data");
        int resultRowNum = 0;

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        ChromeDriver driver = new ChromeDriver(options);
        String headers[] = {"Sl.No.","BSE's Reference No.","Date of receipt","Name of client","Name  of Company","Type of Complaint","Status","Status Date (dd/mm/yyyy)",
                "Date of Filing Arbitration","Name of Arbitrator(s)","Date of Arbitration Award"};

        for(int i=0; i< urls.size(); i++){
            driver.get(urls.get(i));
            String uniqueFileName = urls.get(i);
            File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);

            if (downloadedFile != null) {
                System.out.println("File downloaded successfully: " + downloadedFile.getName());

                // Process zip files to extract
                if (uniqueFileName.contains(".zip")) {
                    String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                    if (extractedFilePath != null) {
                        // Process the extracted Excel file
                        searchAndCopyRow(extractedFilePath, name, downloadDir + "Complaints_received_from_investors_against_Listed_Companies.xlsx", resultSheet,resultRowNum, resultWorkbook, headers);

                        // Clean up: Delete the zip and extracted Excel file after processing
                        Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                        Files.deleteIfExists(Paths.get(extractedFilePath));
                    }
                } else {
                    // Process .xls or .xlsx files directly
                    searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "Complaints_received_from_investors_against_Listed_Companies.xlsx",resultSheet, resultRowNum, resultWorkbook, headers);
                    // Clean up: Delete the Excel file after processing
                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                }
            } else {
                System.out.println("File download failed.");
            }
        }
        driver.quit();

    }


    void compalints_Against_TM(String name, List<String> urls) throws InterruptedException, IOException {
        String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\Complaint_Arbitration_Status\\";

        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Filter Data");
        int resultRowNum = 0;

        // Configure Chrome options for automatic download
        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("download.default_directory", downloadDir);
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        ChromeDriver driver = new ChromeDriver(options);
        String[] headers = {"S.No.","Complaint Ref.no.","Date of Receipt","Name of Client","Type of Complaint","Name of TM #","Market Segment","If defaulter, Date of its declaration","Status",
                "Status Date","Arbitration Ref.no.","Date of Filing Arbitration","Name of Arbitrator (s)","Date of Arbitration Award"};
        for(int i=0; i< urls.size(); i++){
            driver.get(urls.get(i));
            String uniqueFileName = urls.get(i);
            File downloadedFile = waitForFileToDownload(downloadDir, uniqueFileName.contains(".zip") ? ".zip" : (uniqueFileName.contains(".xls") ? ".xls" : ".xlsx"), 30);

            if (downloadedFile != null) {
                System.out.println("File downloaded successfully: " + downloadedFile.getName());

                // Process zip files to extract Annexure.xls or Annexure.xlsx
                if (uniqueFileName.contains(".zip")) {
                    String extractedFilePath = extractZipFile(downloadedFile.getAbsolutePath(), downloadDir);
                    if (extractedFilePath != null) {
                        // Process the extracted Excel file (Annexure.xls or Annexure.xlsx)
                        searchAndCopyRow(extractedFilePath, name, downloadDir + "Complaints_Received_from_clients_against_Trading_Members.xlsx", resultSheet,resultRowNum, resultWorkbook, headers);

                        // Clean up: Delete the zip and extracted Excel file after processing
                        Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                        Files.deleteIfExists(Paths.get(extractedFilePath));
                    }
                } else {
                    // Process .xls or .xlsx files directly
                    searchAndCopyRow(downloadedFile.getAbsolutePath(), name, downloadDir + "Complaints_Received_from_clients_against_Trading_Members.xlsx",resultSheet, resultRowNum, resultWorkbook, headers);
                    // Clean up: Delete the Excel file after processing
                    Files.deleteIfExists(Paths.get(downloadedFile.getAbsolutePath()));
                }
            } else {
                System.out.println("File download failed.");
            }
        }
        driver.quit();

    }
    private static File waitForFileToDownload(String dirPath, String extension, int timeoutInSeconds) throws InterruptedException {
        File dir = new File(dirPath);
        File downloadedFile = null;
        int waited = 0;

        while (waited < timeoutInSeconds) {
            File[] files = dir.listFiles((d, name) -> name.toLowerCase().endsWith(extension));
            if (files != null && files.length > 0) {
                downloadedFile = files[0];
                if (downloadedFile.length() > 0) {  // Ensure file is not empty
                    break;
                }
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

            if (extractedFilePath == null) {
                System.out.println("No File Found in Zip.");
            }
            return extractedFilePath;  // Returns the path to the extracted file or null if not found

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    private void searchAndCopyRow(String excelFilePath, String keyword, String outputFilePath, Sheet resultSheet, int resultRowNum, Workbook  resultWorkbook, String[] headers) {
        // Convert the keyword to lowercase to make the search case-insensitive
        String keywordLowerCase = keyword.toLowerCase();

        File file = new File(excelFilePath);

        // Verify if the file exists and has a valid Excel extension
        if (!file.exists()) {
            System.err.println("File not found: " + excelFilePath);
            return;
        }
        if (!(excelFilePath.endsWith(".xls") || excelFilePath.endsWith(".xlsx"))) {
            System.err.println("Invalid file format: " + excelFilePath);
            return;

        }
        // Create a bold font style for the header row
        CellStyle headerStyle = resultWorkbook.createCellStyle();
        Font boldFont = resultWorkbook.createFont();
        boldFont.setBold(true);
        headerStyle.setFont(boldFont);

        // Set background color for the header row to #C1F0C8
        XSSFColor headerBgColor = new XSSFColor(new byte[]{(byte) 193, (byte) 240, (byte) 200}, null);  // RGB for #C1F0C8
        ((XSSFCellStyle) headerStyle).setFillForegroundColor(headerBgColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row headerRow = resultSheet.createRow(resultRowNum++);
        for(int i=0; i< headers.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
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
            for(int i = 0; i< headers.length; i++){
                resultSheet.autoSizeColumn(i);
            }
            // After copying, save the result workbook to a file
            saveResultWorkbook(outputFilePath, resultWorkbook);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void saveResultWorkbook(String outputFilePath, Workbook  resultWorkbook) {
        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
            resultWorkbook.write(fileOut);
            System.out.println("Filtered data saved to Excel file: " + outputFilePath);
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
class Main3{
    public static void main(String[] args) {
        Complaint_Arbitration_Status obj = new Complaint_Arbitration_Status("Asho");
        Thread t = new Thread(obj);
        t.start();
    }
}
