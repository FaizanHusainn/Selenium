package BSE;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.nio.file.*;
import java.time.Duration;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class Debarred_Entities implements Runnable {

    public String name;

    Debarred_Entities(String name) {
        this.name = name;
    }
    public void run() {
        process();
    }
    public void process() {
        WebDriver driver = null;

        try {
            // Set download directory path
            String downloadDir = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\";
           // String keyword = "Rite"; // Replace with your search keyword

            // Configure Chrome options for download
            ChromeOptions options = new ChromeOptions();
            Map<String, Object> prefs = new HashMap<>();
            prefs.put("download.default_directory", downloadDir);
            prefs.put("download.prompt_for_download", false);
            prefs.put("safebrowsing.enabled", true);
            options.setExperimentalOption("prefs", prefs);
            driver = new ChromeDriver(options);
            driver.manage().window().maximize();

            // Open the BSE URL
            driver.get("https://www.bseindia.com/investors/debent.aspx?expandable=5");

            // Wait for the download link to be clickable and click it
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            String targetText[] = {"Download Archive file-As","Download file-As","Other Competent Authorities Debarred"};
            executeFile(wait,targetText[0],"\\BSE_Debarered_Entities_Results.xlsx", downloadDir);
            executeFile(wait,targetText[1],"\\Entities_debarred_by_SEBI_Results.xlsx", downloadDir);
            executeFile(wait,targetText[2],"\\Entities_debarred_by_other_Competent_Authorities_Results.xlsx", downloadDir);

            driver.get("https://www.bseindia.com/static/investors/list_PAN_deactivated.aspx");
            executeFile(wait,"PAN deactivated on account of non delivery of show cause notice","\\Deactivated_Non-Compliant_Entities_Results.xlsx", downloadDir);



        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }
    public void executeFile(WebDriverWait wait,String targetedText, String fileName, String downloadDir) throws IOException, InterruptedException {
        WebElement downloadLink = wait.until(ExpectedConditions.elementToBeClickable(
                By.partialLinkText(targetedText)));
        downloadLink.click();

        // Wait for the file to download using a more optimized way
        File downloadedZip = waitForFileToDownload(downloadDir, ".zip", 60);  // Wait up to 60 seconds for the download
        if (downloadedZip != null) {
            String extractedFilePath = extractZipFile(downloadedZip.getAbsolutePath(), downloadDir);

            if (extractedFilePath != null) {
                // Search the Excel file and extract the matching row
                searchAndCopyRow(extractedFilePath, name, downloadDir + fileName);

                // Delete the zip and extracted Excel file after processing
                Files.deleteIfExists(Paths.get(downloadedZip.getAbsolutePath()));
                Files.deleteIfExists(Paths.get(extractedFilePath));
                System.out.println("Deleted zip and extracted Excel file.");
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

    // Extract zip file and return the extracted Excel file path
    private static String extractZipFile(String zipFilePath, String extractToDirectory) {
        try (ZipFile zipFile = new ZipFile(zipFilePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();

            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String filePath = extractToDirectory + File.separator + entry.getName();

                if (!entry.isDirectory()) {
                    try (BufferedInputStream bis = new BufferedInputStream(zipFile.getInputStream(entry))) {
                        File newFile = new File(filePath);
                        new File(newFile.getParent()).mkdirs();

                        try (FileOutputStream fos = new FileOutputStream(newFile)) {
                            byte[] bytesIn = new byte[4096];
                            int read;
                            while ((read = bis.read(bytesIn)) != -1) {
                                fos.write(bytesIn, 0, read);
                            }
                        }
                    }
                    System.out.println("Extracted file: " + filePath);
                    return filePath; // Return first extracted file (Assumed Excel file)
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    // Search and copy matching rows from Excel
    private static void searchAndCopyRow(String excelFilePath, String keyword, String outputFilePath) {
        // Convert the keyword to lowercase to make the search case-insensitive
        String keywordLowerCase = keyword.toLowerCase();

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(fis);
             Workbook resultWorkbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.getSheetAt(0);
            Sheet resultSheet = resultWorkbook.createSheet("Filtered Data");

            int resultRowNum = 0;
            CellStyle wrapStyle = resultWorkbook.createCellStyle();
            wrapStyle.setWrapText(true);  // Enable wrap text

            CellStyle headerStyle = resultWorkbook.createCellStyle();
            Font boldFont = resultWorkbook.createFont();
            boldFont.setBold(true);
            headerStyle.setFont(boldFont);

            // Copy header row and make it bold
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                Row resultHeaderRow = resultSheet.createRow(resultRowNum++);
                copyRow(headerRow, resultHeaderRow, headerStyle);
            }

            // Copy the matching rows
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;  // Skip the header row
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        // Compare cell value with keyword, both converted to lowercase
                        String cellValueLowerCase = cell.getStringCellValue().toLowerCase();
                        if (cellValueLowerCase.contains(keywordLowerCase)) {
                            Row resultRow = resultSheet.createRow(resultRowNum++);
                            copyRow(row, resultRow, wrapStyle);
                            break;
                        }
                    }
                }
            }

            // Adjust column width to fit the content and apply wrap
            for (int i = 0; i < resultSheet.getRow(0).getLastCellNum(); i++) {
                resultSheet.autoSizeColumn(i);  // Auto-size the columns
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                resultWorkbook.write(fos);
                System.out.println("Matching rows saved to: " + outputFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Copy one row from source to destination and apply styles
    private static void copyRow(Row sourceRow, Row destinationRow, CellStyle cellStyle) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell destinationCell = destinationRow.createCell(i);

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
                destinationCell.setCellStyle(cellStyle);  // Apply cell style (wrap or bold)
            }
        }
    }
}
class Main1{
    public static void main(String[] args) {
        Debarred_Entities d = new Debarred_Entities("parv");
        Thread t = new Thread(d);
        t.start();
    }
}