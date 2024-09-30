package BSE;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class Defaulter_Members_BSE implements Runnable {

    public String name;

    Defaulter_Members_BSE(String name) {
        this.name = name;
    }

    public void run() {
        try {
            process();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void process() throws IOException {
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));  // Reduced general wait time
        driver.manage().window().maximize();
        driver.get("https://www.bseindia.com/members/List_defaulters_Expelled_members.aspx?expandable=2");
        String directory = "C:\\Users\\Faizan\\Automation\\src\\test\\java\\BSE\\";


        // Prepare to store data in Excel
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Defaulter Members");

        // Create a bold font style for the header row
        CellStyle headerStyle = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        headerStyle.setFont(boldFont);

        // Set background color for the header row to #C1F0C8
        XSSFColor headerBgColor = new XSSFColor(new byte[]{(byte) 193, (byte) 240, (byte) 200}, null);  // RGB for #C1F0C8
        ((XSSFCellStyle) headerStyle).setFillForegroundColor(headerBgColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        // Locate the outer table first
        WebElement outerTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.largetable table")));
        // Now locate the inner table inside the outer table
        WebElement innerTable = outerTable.findElement(By.cssSelector("td table"));  // Adjusted selector to target the inner table
        // Find the tbody of the inner table
        WebElement tbody = innerTable.findElement(By.tagName("tbody"));
        // Get all rows of the inner table
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        WebElement header = rows.get(0);

        List<WebElement> headertds = header.findElements(By.tagName("td"));
        Row row = sheet.createRow(0);
        for( int i=0; i< headertds.size(); i++ ) {
            Cell headerCell = row.createCell(i);
            headerCell.setCellValue(headertds.get(i).getText());
            headerCell.setCellStyle(headerStyle);
        }
        int rowNum = 1;

        for(int i =1 ; i< rows.size(); i++ ) {
            List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
            String memberName = "";
            if(tds.size() <3){
                WebElement previosRow = rows.get(i-1);
                List<WebElement> previostds = previosRow.findElements(By.tagName("td"));
                memberName = previostds.get(2).getText();
            }else{
                memberName = tds.get(2).getText();
            }

            System.out.println("checking for mamber name :" + memberName);
            if(memberName.toLowerCase().contains(name.toLowerCase())) {
                int handleMergecol = 3;
                Row datarow = sheet.createRow(rowNum++);
                for(int j =0; j < tds.size(); j++ ) {
                    if(tds.size() < 3){
                        Cell cell = datarow.createCell(handleMergecol++);
                        cell.setCellValue(tds.get(j).getText());
                    }else{
                        Cell cell = datarow.createCell(j);
                        cell.setCellValue(tds.get(j).getText());
                    }

                }
            }

        }
        for(int i = 0; i< 5; i++){
            sheet.autoSizeColumn(i);
        }


        // Save the Excel file
        FileOutputStream fileOut = new FileOutputStream(directory + "BSE_Defaulter_Member_Results.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();


        driver.quit();  // Close the browser once done
    }
}

class Main {
    public static void main(String[] args) {
        Defaulter_Members_BSE obj = new Defaulter_Members_BSE("Ka");
        Thread thread = new Thread(obj);
        thread.start();
    }
}
