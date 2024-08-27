package National_Investigation_Agency;

import java.time.Duration;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class NIA_Most_Wanted {

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {
            driver.get("https://www.nia.gov.in/most-wanted.htm");

            WebElement search_filter = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label#ContentPlaceHolder1_WantedList1_lblKeyword")));
            search_filter.click();

            WebElement search_Name = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input#ContentPlaceHolder1_WantedList1_txtKeyword")));
            search_Name.sendKeys("Abdul");

            WebElement search_btn = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input#ContentPlaceHolder1_WantedList1_btnSearch")));
            search_btn.click();

            List<WebElement> resultContainers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("ul.commonListing li")));
            System.out.println("Number of result containers found: " + resultContainers.size());

            Set<String> processedNames = new HashSet<>(); // Track processed names
            int counter = 1;

            for (WebElement search_result_li : resultContainers) {
                if (!search_result_li.getText().trim().isEmpty()) {
                    // Ensure the result container is clickable
                    wait.until(ExpectedConditions.elementToBeClickable(search_result_li));
                    
                    // Click on the result and wait for the content to update
                    search_result_li.click();
                    
                    // Wait until the content is updated
                    WebElement nameElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ContentPlaceHolder1_WantedList1_lblName")));
                    String nameText = nameElement.getText().trim();

                    // Ensure the name is not processed multiple times
                    if (!processedNames.contains(nameText)) {
                        System.out.println(counter + ": " + nameText);
                        processedNames.add(nameText);
                        counter++;
                    }

                    // Wait for the result container to become clickable again
                    wait.until(ExpectedConditions.elementToBeClickable(search_result_li));
                    
                    // Add a small delay to ensure content updates
                    Thread.sleep(1000);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
          //  driver.quit(); // Ensure the driver is properly closed
            System.out.println("Execution complete");
        }
    }
}
