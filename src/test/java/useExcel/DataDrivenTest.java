package useExcel;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Created by Diana on 19.01.2017.
 */
public class DataDrivenTest {
    WebDriver driver;
    //Locators
    private By byEmail = By.name("login");
    private By byPassword = By.name("pass");
    private By bySubmit = By.xpath("//*[@id=\"vxod-na-sajt\"]/form/table/tbody/tr[3]/td[2]/input");
    private By byError = By.xpath("");

    @BeforeClass

    public void beforeMethod() throws Exception {

        driver = new ChromeDriver();

        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().window().maximize();

        driver.get("http://pharmbase.com.ua");

    }

    @Test(dataProvider = "Authentication")

    public void getLoginData( List data){


        driver.findElement(byEmail).sendKeys(data.get(0).toString());

        System.out.println(data.get(0).toString());

        driver.findElement(byPassword).sendKeys(data.get(1).toString());

        System.out.println(data.get(1).toString());
        driver.findElement(bySubmit).click();

    }

    @AfterClass

    public void afterMethod() {

        driver.close();

    }

    @DataProvider
    public  Object[][]Authentication() throws Exception {
        DataDrivenExcel userData = new DataDrivenExcel();
        userData.setExcelFile("C:\\Users\\Diana\\IdeaProjects\\group\\src\\test\\java\\testDate\\TestDate.xlsx", "first");
        List rowsNo = userData.getRowContains("user0",0);
        return userData.getTableArray(rowsNo);

    }

}
