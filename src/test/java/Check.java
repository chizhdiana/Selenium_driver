import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.util.concurrent.TimeUnit;

/**
 * Created by Diana on 03.01.2017.
 */
public class Check {
    public WebDriver driver;
    public WebDriverWait wait;
    String appUrl = "http://pharmbase.com.ua/";
    private By byEmail = By.name("login");
    private By byPassword = By.name("pass");
    private By bySubmit = By.xpath("//*[@id=\"vxod-na-sajt\"]/form/table/tbody/tr[3]/td[2]/input");
    private By byError = By.xpath("");
    @BeforeTest
    public void setSetup() {
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        // driver.manage().window().maximize();
        wait = new WebDriverWait(driver, 5);
        driver.get(appUrl);
    }

    @AfterTest
    public void afterTest(){
        driver.quit();
    }

    @Test
    public void startTest(){
        if(  driver.findElement(byEmail).isDisplayed()){
            driver.findElement(byEmail).click();
            driver.findElement(byEmail).sendKeys("name");
        }else {
            System.out.println("not found element ");
        }

    }
}
