
package seleniumwebdriver;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.stereotype.Service;

@Service
public class ScrapingService {

    public String runSeleniumScript(String searchTerm) {
        WebDriver driver = new ChromeDriver();
        String wordFilePath = "C:\\Users\\faiza\\eclipse-workspace\\seleniumwebdriver\\src\\test\\java\\seleniumwebdriver\\Output.docx";

        try {
            driver.get("https://www.nationalsecurity.gov.au/what-australia-is-doing/terrorist-organisations/listed-terrorist-organisations");
            WebElement link = driver.findElement(By.linkText(searchTerm));
            link.click();

            WebElement contentDiv = driver.findElement(By.cssSelector(".content"));
            List<WebElement> elements = contentDiv.findElements(By.xpath(".//*"));

            XWPFDocument document = new XWPFDocument();
            for (WebElement element : elements) {
                String tagName = element.getTagName();
                String text = element.getText().trim();
                if (!text.isEmpty()) {
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();
                    run.setText(text);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(wordFilePath)) {
                document.write(fos);
            }

            document.close();
            return wordFilePath;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        } finally {
            driver.quit();
        }
    }
}
