package National_Investigation_Agency;

import java.time.Duration;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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
        JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;  // Initialize JavaScript Executor

        try {
            driver.get("https://www.nia.gov.in/most-wanted.htm");

            WebElement search_filter = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label#ContentPlaceHolder1_WantedList1_lblKeyword")));
            search_filter.click();

            WebElement search_Name = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input#ContentPlaceHolder1_WantedList1_txtKeyword")));
            search_Name.sendKeys("Abdul");

            WebElement search_btn = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("input#ContentPlaceHolder1_WantedList1_btnSearch")));
            search_btn.click();

            WebElement resultContainers = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='mCSB_1_container']/ul")));
            List<WebElement> li_container = resultContainers.findElements(By.tagName("li"));

            System.out.println("Number of result containers found: " + li_container.size());

            for (int i = 0; i < li_container.size(); i++) {
                WebElement search_result_li = li_container.get(i);
                try {
                    // Scroll the element into view
                    jsExecutor.executeScript("arguments[0].scrollIntoView(true);", search_result_li);

                    // Wait until the element is clickable
                    WebElement clickableElement = wait.until(ExpectedConditions.elementToBeClickable(search_result_li));
                    clickableElement.click();

                    // Add your logic for what to do after clicking the element
//                    System.out.println("Clicked on result item " + i);
                    Thread.sleep(1000);
                    clickableElement.click();
                    System.out.println(i);
                    WebElement Name = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_WantedList1_lblName")));
                    String person_NameString = Name.getText();
                    System.out.println(person_NameString);
                    
                    WebElement wanted_info = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("ul.wantedInfo")));
                    List<WebElement> wanted_info_li = wanted_info.findElements(By.tagName("li"));
                    
                    for (WebElement liElement : wanted_info_li) {
                        // Check if the <li> element is visible
                        String displayStyle = liElement.getCssValue("display");
                        if (!"none".equals(displayStyle)) {
                            List<WebElement> divs = liElement.findElements(By.tagName("div"));
                            for (WebElement div : divs) {
                                if (div.getAttribute("class").equals("leftCol")) {
                                    String leftText = div.findElement(By.tagName("span")).getText();
                                 //   System.out.println("Label: " + leftText);
                                } else if (div.getAttribute("class").equals("rightCol")) {
                                    String rightText = div.findElement(By.tagName("span")).getText();
                                 //   System.out.println("Value: " + rightText);
                                }
                            }
                        }
                    	
                    }
                    
                   
                   


                } catch (Exception e) {
                    System.out.println("Element not clickable or error occurred: " + e.getMessage());
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
           // driver.quit(); // Ensure the driver is properly closed
            System.out.println("Execution complete");
        }
    }
}
