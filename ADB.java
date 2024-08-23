package American_Dev_Bank;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ADB {
	
	 public static void main(String[] args) {
	        WebDriver driver = new ChromeDriver();
	        driver.manage().window().maximize();
	        driver.get("https://lnadbg4.adb.org/oga0009p.nsf/sancALL1P?OpenView&Start=999&Count=999");
	 
	 }

}
