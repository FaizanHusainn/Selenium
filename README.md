https://disc.icai.org/disciplinary-committee-under-section-21d-2/


<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
  <modelVersion>4.0.0</modelVersion>
  <groupId>seleniumwebdriver</groupId>
  <artifactId>seleniumwebdriver</artifactId>
  <version>0.0.1-SNAPSHOT</version>
  
  
  <parent>
    <groupId>org.springframework.boot</groupId>
    <artifactId>spring-boot-starter-parent</artifactId>
    <version>3.1.3</version> <!-- Use the latest stable version here -->
    <relativePath/> <!-- lookup parent from repository -->
</parent>
  
  <dependencies>
<!-- https://mvnrepository.com/artifact/org.seleniumhq.selenium/selenium-java -->
	<dependency>
		 <groupId>org.seleniumhq.selenium</groupId>
		 <artifactId>selenium-java</artifactId>
		 <version>4.23.1</version>
	</dependency>
	
	<dependency>
   <groupId>com.opencsv</groupId>
   <artifactId>opencsv</artifactId>
   <version>5.5.2</version>
  </dependency>
   
   <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>
        
        <!-- Thymeleaf Template Engine -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-thymeleaf</artifactId>
        </dependency>

<dependency>

   <groupId>javax.servlet</groupId>

   <artifactId>javax.servlet-api</artifactId>

   <version>4.0.1</version>

   <scope>provided</scope>

</dependency>
 <!-- Apache POI for working with Word files -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.3</version>
    </dependency>
    
    

    <!-- Apache POI dependencies for XML schemas and formats -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml-schemas</artifactId>
        <version>4.1.2</version>
    </dependency>
    
    <!-- XMLBeans (used by POI) -->
    <dependency>
        <groupId>org.apache.xmlbeans</groupId>
        <artifactId>xmlbeans</artifactId>
        <version>5.1.1</version>
    </dependency>

    <!-- Commons Collections (used by POI) -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-collections4</artifactId>
        <version>4.4</version>
    </dependency>
    
      <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>

    <!-- Spring Boot DevTools (Optional, for development) -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-devtools</artifactId>
        <scope>runtime</scope>
        <optional>true</optional>
    </dependency>
    
      <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-thymeleaf</artifactId>
    </dependency>
    
</dependencies>
 <build>
    <plugins>
        <plugin>
            <groupId>org.apache.maven.plugins</groupId>
            <artifactId>maven-compiler-plugin</artifactId>
            <version>3.8.0</version>
            <configuration>
                <source>8</source>
                <target>8</target>
            </configuration>
        </plugin>
        <plugin>
            <groupId>org.apache.maven.plugins</groupId>
            <artifactId>maven-surefire-plugin</artifactId>
            <version>3.0.0-M1</version>
            <configuration>
                <useSystemClassLoader>false</useSystemClassLoader>
                <forkCount>1</forkCount>
                <useFile>false</useFile>
                <skipTests>false</skipTests>
                <testFailureIgnore>true</testFailureIgnore>
                <forkMode>once</forkMode>
                <suiteXmlFiles>
                    <suiteXmlFile>testing.xml</suiteXmlFile>
                </suiteXmlFiles>
            </configuration>
        </plugin>
    </plugins>
</build>
<properties>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
</properties>
</project>




 -------------------------------------------------------------------------------

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

