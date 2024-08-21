import java.io.File;
import java.net.MalformedURLException;
import java.time.Duration;
import java.util.Base64;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;


import org.json.JSONException;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.anti_captcha.Api.ImageToText;
import com.anti_captcha.Helper.DebugHelper;
//import com.google.common.io.Files;


public class DRT {
    public static void main(String[] args)  throws InterruptedException, MalformedURLException, JSONException {
//    	Scanner sc = new Scanner(System.in);
//    	System.out.print("Enter Party Name : ");
//        String party_Name = sc.nextLine();
         
         
        WebDriver driver = new ChromeDriver();
        

        try {
            String targetUrl = "https://drt.gov.in/#/casedetail";
            
            // Navigate to the URL
            driver.get(targetUrl);
            driver.manage().window().maximize();

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
            System.out.println("Wait");
            
//            driver.findElement(By.linkText("Case Details")).click();
            
            // Wait until the "DRAT Case Status" link is visible
            WebElement partyName = wait.until(
                ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div/div[2]/div[3]/div/div/div[1]/ul/li[3]/span")));
            partyName.click();
                
            WebElement dropdownElement = wait.until(
            ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"schemeNameDratDrtId\"]"))); // Use the correct locator

                // Initialize the Select object
            Select dropdown = new Select(dropdownElement);

                // Select an option by visible text
            String visibleText = "DEBT RECOVERY APPELLATE TRIBUNAL - DELHI"; // Replace with the visible text you want to select
            dropdown.selectByVisibleText(visibleText);
            
            
            WebElement select = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"partyName\"]")));
            select.sendKeys("abc");
            
//            JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
//            String base64Image = (String) jsExecutor.executeScript(
//                    "return document.querySelector('.rnc-canvas').toDataURL('image/png').substring(22);"
//                );
//            
//            
//            System.out.println(base64Image);
            // Find the canvas element
            WebElement canvas = driver.findElement(By.cssSelector("canvas.rnc-canvas"));

            // Take a screenshot of the canvas element
            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

            // Read the screenshot file and encode it to base64
            byte[] fileContent = Files.readAllBytes(screenshot.toPath());
            String base64Image = Base64.getEncoder().encodeToString(fileContent);

            // Output the base64 string
            System.out.println("Base64 Image String: " + base64Image);
            
            //Solving captcha
            DebugHelper.setVerboseMode(true);

            ImageToText api = new ImageToText();
            api.setClientKey("7ce431ab54f911639271322ddd8ddb6e");
            api.setFilePath(base64Image);

            //Specify softId to earn 10% commission with your app.
            //Get your softId here: https://anti-captcha.com/clients/tools/devcenter
            api.setSoftId(0);

            //optional parameters, see documentation
            //api.setPhrase(true);                                    // 2 words
            //api.setCase(true);                                      // case sensitivity
            //api.setNumeric(ImageToText.NumericOption.NUMBERS_ONLY); // only numbers
            //api.setMath(1);                                         // do math operation like result of 50+5
            //api.setMinLength(1);                                    // minimum length of the captcha text
            //api.setMaxLength(10);                                   // maximum length of the captcha text
            //api.setLanguagePool("en");                              // set pool of the languages used in the captcha

            if (!api.createTask()) {
                DebugHelper.out(
                        "API v2 send failed. " + api.getErrorMessage(),
                        DebugHelper.Type.ERROR
                );
            } else if (!api.waitForResult()) {
                DebugHelper.out("Could not solve the captcha.", DebugHelper.Type.ERROR);
            } else {
                DebugHelper.out("Result: " + api.getTaskSolution().getText(), DebugHelper.Type.SUCCESS);
            }
            
            
            WebElement captcha_input = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div/div[2]/div[3]/div/div/div[1]/div/form/div/div[3]/div[1]/div/input")));
            captcha_input.sendKeys("1234");
               
            

        } catch (Exception e) {
            e.printStackTrace(); // Print the stack trace for debugging
        } finally {
            // Close the browser
           // driver.quit();
        }
    }
}


--------------------------------------------------------

    import java.io.File;
import java.net.MalformedURLException;
import java.time.Duration;
import java.util.Base64;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.json.JSONException;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.anti_captcha.Api.ImageToText;
import com.anti_captcha.Helper.DebugHelper;
import com.anti_captcha.Api.CreateTask;
import com.anti_captcha.ApiResult.ImageToTextResult;

public class DRT {
    public static void main(String[] args) throws InterruptedException, MalformedURLException, JSONException {
        WebDriver driver = new ChromeDriver();
        
        try {
            String targetUrl = "https://drt.gov.in/#/casedetail";
            
            // Navigate to the URL
            driver.get(targetUrl);
            driver.manage().window().maximize();

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
            System.out.println("Wait");

            WebElement partyName = wait.until(
                ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div/div[2]/div[3]/div/div/div[1]/ul/li[3]/span")));
            partyName.click();
                
            WebElement dropdownElement = wait.until(
                ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"schemeNameDratDrtId\"]"))); 

            Select dropdown = new Select(dropdownElement);
            String visibleText = "DEBT RECOVERY APPELLATE TRIBUNAL - DELHI";
            dropdown.selectByVisibleText(visibleText);
            
            WebElement select = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"partyName\"]")));
            select.sendKeys("abc");
            
            // Solve CAPTCHA
            WebElement captchaImage = driver.findElement(By.xpath("//img[@alt='Captcha Image']")); // Locate the CAPTCHA image
            String captchaImageBase64 = captchaImage.getScreenshotAs(OutputType.BASE64); // Capture the image as Base64

            // Save the CAPTCHA image temporarily to solve it using Anti-Captcha
            Path captchaPath = Paths.get("captcha.png");
            byte[] decodedBytes = Base64.getDecoder().decode(captchaImageBase64);
            Files.write(captchaPath, decodedBytes);

            String apiKey = "your_anti_captcha_api_key"; // Replace with your Anti-Captcha API key

            ImageToText api = new ImageToText();
            api.setClientKey(apiKey);
            api.setFilePath(captchaPath.toString());

            if (!api.createTask()) {
                System.out.println("API v2 send failed. " + api.getErrorMessage());
                return;
            }

            if (!api.waitForResult()) {
                System.out.println("Could not solve the captcha.");
                return;
            }

            System.out.println("Captcha Text: " + api.getTaskSolution().getText());

            // Fill in the solved CAPTCHA text
            WebElement captchaInput = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div/div/div[2]/div[3]/div/div/div[1]/div/form/div/div[3]/div[1]/div/input")));
            captchaInput.sendKeys(api.getTaskSolution().getText());
               
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }
}




