package AFDB;

import java.io.FileOutputStream;
import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class AFDB {

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.afdb.org/en/projects-operations/debarment-and-sanctions-procedures");

        try {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
            WebElement search_Name = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"datatable-1_filter\"]/label/input")));
            search_Name.sendKeys("Af");
             
            WebElement total_entries = driver.findElement(By.xpath("//*[@id=\"datatable-1_info\"]"));
            String text =  total_entries.getText();  
            // Regular expression to find the number after "of"
            String regex = "of\\s([\\d,]+)";
            
            // Match the pattern in the string
            java.util.regex.Pattern pattern = java.util.regex.Pattern.compile(regex);
            java.util.regex.Matcher matcher = pattern.matcher(text);
            
            int Total_search_result =0;

            if (matcher.find()) {
                // Extract the matched number (with commas)
                String numberWithCommas = matcher.group(1);

                // Remove the commas
                String numberWithoutCommas = numberWithCommas.replace(",", "");

                // Convert the number to an integer
                Total_search_result= Integer.parseInt(numberWithoutCommas);

                // Output the result
                System.out.println("Total Search result: " + Total_search_result);
            } else {
                System.out.println("No match found.");
            }
            
            
            // Create a new workbook and sheet for Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("AFDB Data");

            // Create a bold font for the headers
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // Initialize row number for Excel
            int count= 0;
            int rowNum = 0;

            while (count < Total_search_result) {
                WebElement table = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("table#datatable-1")));

                // Extract table headers
                List<WebElement> headers = table.findElements(By.cssSelector("thead th"));

                // Write the header row to the Excel sheet if it's the first page
                if (rowNum == 0) {
                    Row headerRow = sheet.createRow(rowNum++);
                    for (int i = 0; i < headers.size(); i++) {
                        Cell cell = headerRow.createCell(i);
                        cell.setCellValue(headers.get(i).getText());
                        cell.setCellStyle(headerCellStyle);
                    }
                }

                // Extract table rows
                List<WebElement> rows = table.findElements(By.cssSelector("tbody tr"));
                

                // Write the table data to the Excel sheet
                for (int i = 0; i < rows.size(); i++) {
                    Row row = sheet.createRow(rowNum++);
                    List<WebElement> cells = rows.get(i).findElements(By.tagName("td"));
                    for (int j = 0; j < cells.size(); j++) {
                        row.createCell(j).setCellValue(cells.get(j).getText());
                        System.out.println(cells.get(j).getText());
                    }
                    count++;
                }

                // Resize all columns to fit the content size
                for (int i = 0; i < headers.size(); i++) {
                    sheet.autoSizeColumn(i);
                }

                // Locate the "Next" button on the current page
                WebElement nextButton = driver.findElement(By.cssSelector("a#datatable-1_next"));

                // Check if the "Next" button is clickable
                if (nextButton != null && nextButton.isEnabled()) {
                    // Click the "Next" button to go to the next page
                    nextButton.click();

                    // Wait for the next page's table to load
//                    wait.until(ExpectedConditions.stalenessOf(table));
                } else {
                    // Break the loop if the "Next" button is not clickable
                    System.out.println("Next button not active");
                    break;
                    
                }
               
            }
            
            String filePath = "C:\\Users\\FaizanHusain\\eclipse-workspace\\Automation\\src\\test\\java\\AFDB\\AFDB_Data.xlsx";

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            // Close the workbook
            workbook.close();

            System.out.println("Data successfully written to Excel file.");

        } catch (Exception e) {
            e.printStackTrace(); // Print the stack trace for debugging
        } finally {
            // Close the browser
            driver.quit();
            System.out.println("Execution complete");
        }
    }
}
