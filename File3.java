package seleniumwebdriver;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class File3 {
    public static void main(String[] args) {
        // Initialize the WebDriver
        Scanner sc = new Scanner(System.in);
        System.out.print("Enter the search name: ");
        String headingSearch = sc.nextLine();
        sc.close();

        WebDriver driver = new ChromeDriver();

        // Define the output Word file path
        String wordFilePath = "C:\\Users\\faiza\\eclipse-workspace\\seleniumwebdriver\\src\\test\\java\\seleniumwebdriver\\Output.docx";

        try {
            // Maximize the browser window
            driver.manage().window().maximize();

            // Navigate to the target website
            driver.get("https://www.nationalsecurity.gov.au/what-australia-is-doing/terrorist-organisations/listed-terrorist-organisations");
            
            // Click the link to navigate to the section
            WebElement link = driver.findElement(By.linkText(headingSearch));
            link.click();

            // Locate the parent content div by class
            WebElement contentDiv = driver.findElement(By.cssSelector(".content"));

            // Create a new Word document
            XWPFDocument document = new XWPFDocument();

            // Use JavaScript to get all elements inside the content div
            JavascriptExecutor js = (JavascriptExecutor) driver;
            List<WebElement> childElements = (List<WebElement>) js.executeScript(
                "let elements = []; " +
                "let allElements = arguments[0].getElementsByTagName('*'); " +
                "for (let i = 0; i < allElements.length; i++) { " +
                "  if (allElements[i].innerText.trim() !== '' && allElements[i].children.length === 0) { " +
                "    elements.push(allElements[i]); " +
                "  } " +
                "} " +
                "return elements;", contentDiv);

            // Extract and format the content
            extractContent(childElements, document);

            // Save the document to a file
            try (FileOutputStream fos = new FileOutputStream(wordFilePath)) {
                document.write(fos);
            }
            document.close();

            System.out.println("Document saved successfully to " + wordFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close the browser
            driver.quit();
        }
    }

    private static void extractContent(List<WebElement> elements, XWPFDocument document) {
        for (WebElement childElement : elements) {
            String tagName = childElement.getTagName().toLowerCase();
            String text = childElement.getText().trim();
            

            if (!text.isEmpty()) { // Only process elements with non-empty text
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();

                switch (tagName) {
                    case "h1":
                        run.setBold(true);
                        run.setFontSize(20);
                        break;
                    case "h2":
                        run.setBold(true);
                        run.setFontSize(16);
                        break;
                    case "h3":
                        run.setBold(true);
                        run.setFontSize(14);
                        break;
                    case "p":
                        run.setFontSize(12);
                        break;
                    case "span":
                        run.setBold(true);
                        run.setFontSize(14);
                        break;
                    case "ul":
                    case "ol":
                        // Handle unordered and ordered lists
                        List<WebElement> listItems = childElement.findElements(By.tagName("li"));
                        for (WebElement listItem : listItems) {
                            XWPFParagraph listParagraph = document.createParagraph();
                            XWPFRun listItemRun = listParagraph.createRun();
                            listItemRun.setText("• " + listItem.getText().trim());
                        }
                        continue;
                    default:
                        run.setFontSize(12);
                        break;
                }

                run.setText(text);
            }
        }
    }
}
