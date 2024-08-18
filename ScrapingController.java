package seleniumwebdriver;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class ScrapingController {

    @Autowired
    private ScrapingService scrapingService;

    @PostMapping("/scrape")
    @ResponseBody
    public String scrape(@RequestParam String searchTerm, Model model) {
        String filePath = scrapingService.runSeleniumScript(searchTerm);
        return "Data has been scraped and saved to: " + filePath;
    }
}
