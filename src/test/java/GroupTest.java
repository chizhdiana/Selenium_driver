import org.openqa.selenium.By;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.util.concurrent.TimeUnit;

/**
 * Created by Diana on 06.12.2016.
 */
public class GroupTest {

    public static WebDriver driver;


    @Test
    public void main() {
        //WebElement element = driver.findElement(By.id("SubmitButton"));
        WebElement element = driver.findElement(By.xpath("//*[@id=\"conteiner\"]/div/noindex[1]/div/a/img"));
        Point point = element.getLocation();
        System.out.println("X cordinate : " + point.x + "Y cordinate: " + point.y);
        String linkText = element.getText();
        System.out.println("Text" + linkText);

        if (element.isDisplayed()) {
            element.click();
        }


        String actualUrl = " http://www.apteka.ua/";
        String currentUrll = driver.getCurrentUrl();
        if (currentUrll.equals(actualUrl)) {
            System.out.println("Verification Successful - The correct Url is opened.");
        } else {
            System.out.println("Actual URL is : " + actualUrl);
            System.out.println("Expected URL is : " + currentUrll);
        }
        driver.navigate().back();
    }

    @BeforeMethod
    public void beforeMethod() {



        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        driver.get("http://www.apteka.ua/pharmzakaz");
    }

    @AfterMethod
    public void afterMethod() {
        driver.close();
    }
}
//*[@id="conteiner"]/div/div[1]/center/iframe 1   <iframe src="http://ad.adriver.ru/cgi-bin/erle.cgi?sid=170048&amp;target=blank&amp;w=550&amp;h=100&amp;bt=51&amp;pz=0&amp;rnd=787833088&amp;tail256=http%3A//pda.apteka.ua/pharmzakaz" frameborder="0" vspace="0" hspace="0" width="550" height="100" marginwidth="0" marginheight="0" scrolling="no"></iframe>
//*[@id="conteiner"]/div/div[1]/center/iframe 2
//<div is="gwd-pagedeck" class="gwd-page-container gwd-pagedeck" id="pagedeck">
//<div is="gwd-page" id="page1" class="gwd-page-wrapper gwd-page-size gwd-lightbox gwd-div-lc4o gwd-page gwd-inactive" data-gwd-width="550px" data-gwd-height="100px" style="">
//<div class="gwd-page-content gwd-page-size gwd-div-17cg">
//<img is="gwd-image" source="bg.jpg" id="gwd-image_3" class="gwd-img-1vqv gwd-img-b723" src="bg.jpg">
//<img is="gwd-image" source="book.png" id="gwd-image_1" class="gwd-img-9ehx gwd-img-coxc gwd-gen-x2fdgwdanimation" style="" src="book.png">
//<img is="gwd-image" source="label1.png" id="gwd-image_2" class="gwd-img-kvjf gwd-gen-r259gwdanimation" style="" src="label1.png">
//<div class="gwd-animation-event event-2-animation" data-event-name="event-2" data-event-time="3000"></div>
//<gwd-taparea id="gwd-taparea_2" class="gwd-taparea-8sm1 gwd-taparea-1oi9"></gwd-taparea>
//</div>
//</div>
//<div id="page1_1" is="gwd-page" class="gwd-page-wrapper gwd-lightbox gwd-div-1k78 gwd-page gwd-play-animation" data-gwd-width="550px" data-gwd-height="100px" style="">
//<div class="gwd-page-content gwd-div-dgck">
//<img is="gwd-image" source="bg.jpg" id="gwd-image_7" class="gwd-img-1vqv gwd-img-kl6f" src="bg.jpg">
//<img is="gwd-image" source="book.png" id="gwd-image_4" class="gwd-img-9ehx gwd-img-1ugu" src="book.png">
//<img is="gwd-image" source="label2.png" id="gwd-image_5" class="gwd-img-1867 gwd-gen-5xx2gwdanimation" style="" src="label2.png">
//<img is="gwd-image" source="lable3.png" id="gwd-image_6" class="gwd-img-1cvn gwd-gen-1bvzgwdanimation" src="lable3.png">
//<div class="gwd-animation-event event-1-animation" data-event-name="event-1" data-event-time="3500"></div>
//<gwd-taparea id="gwd-taparea_1" class="gwd-taparea-8sm1 gwd-taparea-1mk2"></gwd-taparea>
//</div>
//</div>
//</div>