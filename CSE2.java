package Calcutta_Stock_Exchange;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionURI;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationLink;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.adobe.pdfservices.operation.PDFServices;
import com.adobe.pdfservices.operation.PDFServicesMediaType;
import com.adobe.pdfservices.operation.PDFServicesResponse;
import com.adobe.pdfservices.operation.auth.Credentials;
import com.adobe.pdfservices.operation.auth.ServicePrincipalCredentials;
import com.adobe.pdfservices.operation.exception.SDKException;
import com.adobe.pdfservices.operation.exception.ServiceApiException;
import com.adobe.pdfservices.operation.exception.ServiceUsageException;
import com.adobe.pdfservices.operation.io.Asset;
import com.adobe.pdfservices.operation.io.StreamAsset;
import com.adobe.pdfservices.operation.pdfjobs.jobs.OCRJob;
import com.adobe.pdfservices.operation.pdfjobs.result.OCRResult;
import com.google.api.client.util.IOUtils;

public class CSE2 {
    
    // Initialize the logger.
    private static final Logger LOGGER = LoggerFactory.getLogger(OcrPDF.class);

    // Method to read and extract text from a PDF from a URL
    public String ReadPDFfile(String pdfUrl) throws Exception {
        // Step 1: Open a stream to the PDF URL (no file is downloaded to disk)
        try (InputStream inputStream = new URL(pdfUrl).openStream()) {
            // Step 2: Load the PDF document directly from the InputStream
            PDDocument pdfDocument = PDDocument.load(inputStream);

            // Step 3: Get the number of pages (optional, just for info)
            System.out.println("Number of pages in the document: " + pdfDocument.getPages().getCount());

            // Step 4: Extract the text from the PDF
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String doctextString = pdfStripper.getText(pdfDocument);

            // Step 5: Print the extracted text
            System.out.println("Extracted text from the document:");

            // Close the PDF document
            pdfDocument.close();

            return doctextString;
        }
    }

    // Method to extract URLs from the PDF using annotations
    public static List<String> extractUrlsFromPDFUsingAnnotations(String pdfPath) {
        List<String> urls = new ArrayList<>();
        try (PDDocument document = PDDocument.load(new File(pdfPath))) {
            for (PDPage page : document.getPages()) {
                for (PDAnnotation annotation : page.getAnnotations()) {
                    if (annotation instanceof PDAnnotationLink) {
                        PDAnnotationLink link = (PDAnnotationLink) annotation;
                        if (link.getAction() instanceof PDActionURI) {
                            PDActionURI uri = (PDActionURI) link.getAction();
                            String url = uri.getURI();
                            urls.add(url);
                        }
                    }
                }
            }

            if (urls.isEmpty()) {
                System.out.println("No URLs found in the PDF using annotations.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return urls;
    }

    // Function to download the main PDF
    public static void downloadPDF(String pdfUrl, String pdfPath) {
        try (InputStream in = new URL(pdfUrl).openStream()) {
            Files.copy(in, Paths.get(pdfPath));
            System.out.println("PDF downloaded successfully to " + pdfPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to write the extracted text to a Word file with proper formatting
    public static void writeTextToWordFile(String text, String wordFilePath) {
        try (XWPFDocument document = new XWPFDocument(); FileOutputStream out = new FileOutputStream(new File(wordFilePath))) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();

            // Split the text by newlines and add each line to the Word document properly
            String[] lines = text.split("\n");
            for (String line : lines) {
                run.setText(line);
                run.addBreak(); // Add a line break after each line
            }

            // Write the document to a file
            document.write(out);
            System.out.println("Text written to Word file: " + wordFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    // Converting the PDF to OCR
    public static synchronized void convertPdfToOCR(String pdfURL) throws Exception {
        try (InputStream inputStream = new URL(pdfURL).openStream()) {
             // Initial setup, create credentials instance
            Credentials credentials = new ServicePrincipalCredentials("5e9dc36b1ad641ec93e28886639405e4","p8e-Nowt1IATwFszgEAaC4GK8p4FpRxT10xr");

            // Creates a PDF Services instance
            PDFServices pdfServices = new PDFServices(credentials);

            // Creates an asset(s) from source file(s) and upload
            Asset asset = pdfServices.upload(inputStream, PDFServicesMediaType.PDF.getMediaType());

            // Creates a new job instance
            OCRJob ocrJob = new OCRJob(asset);

            // Submit the job and gets the job result
            String location = pdfServices.submit(ocrJob);
            PDFServicesResponse<OCRResult> pdfServicesResponse = pdfServices.getJobResult(location, OCRResult.class);

            // Get content from the resulting asset(s)
            Asset resultAsset = pdfServicesResponse.getResult().getAsset();
            StreamAsset streamAsset = pdfServices.getContent(resultAsset);

            // Creates an output stream and copy stream asset's content to it
            Path outputDir = Paths.get("C:\\Users\\FaizanHusain\\eclipse-workspace\\Automation\\src\\test\\java\\Calcutta_Stock_Exchange");
            if (!Files.exists(outputDir)) {
                Files.createDirectories(outputDir);  // Create directory if it doesn't exist
            }

            // Output file path
            Path outputPath = outputDir.resolve("ocrOutput.pdf");

            // Creates an output stream and copy stream asset's content to it
            try (OutputStream outputStream = Files.newOutputStream(outputPath)) {
                LOGGER.info("Saving asset at " + outputPath.toString());
                IOUtils.copy(streamAsset.getInputStream(), outputStream);  // Copy the stream content
            }
       } catch (ServiceApiException | IOException | SDKException | ServiceUsageException ex) {
        LOGGER.error("Exception encountered while executing operation", ex);
       }
  }

    public String ReadOCRPDFfile(String pdfUrl) throws Exception {
        try (InputStream inputStream = Files.newInputStream(new File(pdfUrl).toPath())) {
            PDDocument pdfDocument = PDDocument.load(inputStream);
            System.out.println("Number of pages in the document: " + pdfDocument.getPages().getCount());
            PDFTextStripper pdfStripper = new PDFTextStripper();
            String doctextString = pdfStripper.getText(pdfDocument);
            pdfDocument.close();
            return doctextString;
        }
    }
    public static void capturePDFScreenshotsToWord(String pdfUrl, String wordFilePath) throws Exception {
        try (InputStream inputStream = new URL(pdfUrl).openStream()) {
            // Load the PDF document
            PDDocument pdfDocument = PDDocument.load(inputStream);
            PDFRenderer pdfRenderer = new PDFRenderer(pdfDocument);
            PDPageTree pages = pdfDocument.getPages();

            // Create a new Word document
            try (XWPFDocument wordDocument = new XWPFDocument(); FileOutputStream out = new FileOutputStream(new File(wordFilePath))) {
                
                // Loop through all the pages of the PDF
                for (int i = 0; i < pages.getCount(); i++) {
                    // Render the page as an image
                    BufferedImage image = pdfRenderer.renderImageWithDPI(i, 300);  // 300 DPI for high-quality image
                    String imagePath = "screenshot_page_" + (i + 1) + ".png";
                    
                    // Save the image to the disk (temporary file)
                    File outputfile = new File(imagePath);
                    ImageIO.write(image, "png", outputfile);
                    
                    // Insert the image into the Word document
                    XWPFParagraph paragraph = wordDocument.createParagraph();
                    XWPFRun run = paragraph.createRun();
                    run.addPicture(Files.newInputStream(Paths.get(imagePath)),
                                   XWPFDocument.PICTURE_TYPE_PNG, imagePath, Units.toEMU(500), Units.toEMU(700));  // Add the image
                }

                // Write the Word document to a file
                wordDocument.write(out);
                System.out.println("Screenshots saved to Word file: " + wordFilePath);
            }

            pdfDocument.close();  // Close the PDF document
        }
    }
    public static void capturePDFScreenshotsToWord(String pdfUrl, XWPFDocument wordDocument) throws Exception {
        PDDocument pdfDocument = null;  // Initialize outside try block
        InputStream inputStream = null;  // Initialize outside try block
        try {
            // Open the URL stream and load the PDF document
            inputStream = new URL(pdfUrl).openStream();
            pdfDocument = PDDocument.load(inputStream);
            PDFRenderer pdfRenderer = new PDFRenderer(pdfDocument);
            PDPageTree pages = pdfDocument.getPages();

            // Add a page break to start a new page for each URL and PDF screenshots
            XWPFParagraph pageBreakParagraph = wordDocument.createParagraph();
            XWPFRun pageBreakRun = pageBreakParagraph.createRun();
            pageBreakRun.addBreak(BreakType.PAGE);  // Start on a new page

            // Add the URL as a heading in the Word document (no extra breaks)
            XWPFParagraph urlParagraph = wordDocument.createParagraph();
            XWPFRun urlRun = urlParagraph.createRun();
            urlRun.setText("URL: " + pdfUrl);  // Add the URL text
            urlRun.setBold(true);  // Make the URL bold for emphasis
            urlRun.addBreak();  // Add a single line break after the URL

            // Loop through all the pages of the PDF
            for (int i = 0; i < pages.getCount(); i++) {
                // Render the page as an image (reduced size for better fit)
                BufferedImage image = pdfRenderer.renderImageWithDPI(i, 150);  // Reduce DPI to 150 for smaller size
                String imagePath = "screenshot_page_" + (i + 1) + "_" + UUID.randomUUID() + ".png";  // Unique image name

                // Save the image to the disk (temporary file)
                File outputfile = new File(imagePath);
                ImageIO.write(image, "png", outputfile);

                // Insert the image into the Word document (adjust size)
                XWPFParagraph imageParagraph = wordDocument.createParagraph();
                XWPFRun run = imageParagraph.createRun();
                run.addPicture(Files.newInputStream(Paths.get(imagePath)),
                               XWPFDocument.PICTURE_TYPE_PNG, imagePath, Units.toEMU(400), Units.toEMU(550));  // Smaller image size
            }
        } finally {
            // Ensure that the PDF document and input stream are closed
            if (pdfDocument != null) {
                pdfDocument.close();
            }
            if (inputStream != null) {
                inputStream.close();
            }
        }
    }



    public static void main(String[] args) throws Exception {
        // Initialize necessary resources
        ExecutorService executor = null;
        try {
            CSE2 obj = new CSE2();

            // Define the URL and file paths
            String pdfUrl = "https://www.cse-india.com/upload/upload/Notice_to_Non-compliant_Listed_Companies1.pdf";
            String pdfPath = "C:\\Users\\FaizanHusain\\eclipse-workspace\\Automation\\src\\test\\java\\Calcutta_Stock_Exchange\\Notice_to_Non-compliant_Listed_Companies1.pdf";
            String wordFilePath = "C:\\Users\\FaizanHusain\\eclipse-workspace\\Automation\\src\\test\\java\\Calcutta_Stock_Exchange\\output.docx";
            String searchKeyword = "Basant";

            // Step 2: Download the main PDF and extract URLs from annotations
            downloadPDF(pdfUrl, pdfPath);
            List<String> urls = extractUrlsFromPDFUsingAnnotations(pdfPath);

            // Create an ExecutorService with a fixed thread pool
            executor = Executors.newFixedThreadPool(5);  // Adjust thread pool size based on your CPU cores
            List<Future<String>> futures = new ArrayList<>();
            List<String> matchingUrls = new ArrayList<>();  // Collect URLs with the keyword
         //   int count = 0;
            // Step 3: Process each extracted URL in parallel
            for (String url : urls) {
//                if (count == 10) break;
//                count++;
                futures.add(executor.submit(() -> {
                    String pageContent = obj.ReadPDFfile(url);

                    if (pageContent.trim().length()==0) {
                        System.out.println("Scanner Document Detected");
                        convertPdfToOCR(url);
                        String OCR_converted_pdf = "C:\\Users\\FaizanHusain\\eclipse-workspace\\Automation\\src\\test\\java\\Calcutta_Stock_Exchange\\ocrOutput.pdf";
                        pageContent = obj.ReadOCRPDFfile(OCR_converted_pdf);
                    }

                    // Check if the search keyword is found in the PDF content
                    if (pageContent.toLowerCase().contains(searchKeyword.toLowerCase())) {
                        System.out.println("Keyword found on page: " + url);
                        return url;  // Return the matching URL
                    }

                    return "";  // Return empty if no match
                }));
            }

            // Step 4: Collect results (the URLs that match the keyword)
            for (Future<String> future : futures) {
                try {
                    String result = future.get();  // Wait for each future task to complete
                    if (!result.isEmpty()) {
                        matchingUrls.add(result);  // Add the matching URL
                    }
                } catch (Exception e) {
                    e.printStackTrace();  // Handle any exceptions from the future tasks
                }
            }

            // Step 5: Process matching URLs sequentially and add screenshots to the Word document
            try (XWPFDocument wordDocument = new XWPFDocument(); FileOutputStream out = new FileOutputStream(new File(wordFilePath))) {
                for (String matchingUrl : matchingUrls) {
                    capturePDFScreenshotsToWord(matchingUrl, wordDocument);  // Add screenshots and URLs sequentially
                }

                // Write the accumulated screenshots and URLs to the Word file
                wordDocument.write(out);  // Save the accumulated content to the Word file
                System.out.println("All matching PDF screenshots saved to Word file: " + wordFilePath);
            } catch (Exception e) {
                e.printStackTrace();  // Handle any exceptions while processing
            }

        } finally {
            // Always shut down the executor
            if (executor != null) {
                executor.shutdown();
            }
        }
    }



}

