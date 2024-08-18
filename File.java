package seleniumwebdriver;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class File {
    public static void main(String args[]) {
        // Create a new instance of ChromeOptions
        ChromeOptions options = new ChromeOptions();
        
        // Add the headless argument
        options.addArguments("--headless");
        
        // Initialize ChromeDriver with the options
        ChromeDriver driver = new ChromeDriver(options);
        
        // Open the URL
        driver.get("https://demo.opencart.com/");
        
        // Get the title of the page
        String title = driver.getTitle();
        
        // Check if the title matches
        if (title.equals("Your Store")) {
            System.out.println("Title matched");
        } else {
            System.out.println("Not Matched");
        }
        
        // Close the driver
        driver.quit();
    }
}
