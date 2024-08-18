package seleniumwebdriver;

import java.util.Scanner;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ScrapeOrderedText {
    public static void main(String[] args) {
        // Set up the WebDriver (ChromeDriver in this case)

        Scanner sc = new Scanner(System.in);
        System.out.print("Enter the search name: ");
        String headingSearch = sc.nextLine();
        sc.close();

        // Initialize the WebDriver

        WebDriver driver = new ChromeDriver();
        try {

            // Maximize the browser window

            driver.manage().window().maximize();

            driver.get("https://www.nationalsecurity.gov.au/what-australia-is-doing/terrorist-organisations/listed-terrorist-organisations");

            // Click the link to navigate to the section
            WebElement link = driver.findElement(By.linkText(headingSearch));

            link.click();
            // Locate the main content container (use correct selector based on inspection)

            WebElement container = driver.findElement(By.cssSelector(".content"));
            
            System.out.print( container.getText());

     
    }
        catch (Exception e) {

            e.printStackTrace();

        } finally {

            // Close the browser

            driver.quit();

        }
    }
}
