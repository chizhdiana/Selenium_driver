package useExcel;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.concurrent.TimeUnit;

/**
 * Created by Diana on 11.01.2017.
 */
public class DataProviderWithExcel_002 {

  //  private int iTestCaseRow;

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

    public void getLoginData( String sUserName, String sPassword) {

        driver.findElement(byEmail).sendKeys(sUserName.toString());

        System.out.println(sUserName);

        driver.findElement(byPassword).sendKeys(sPassword.toString());

        System.out.println(sPassword);
        driver.findElement(bySubmit).click();

    }

    @AfterClass

    public void afterMethod() {

        driver.close();

    }

    @DataProvider

    public Object[][] Authentication() throws Exception{

        useExcel.ExcelUtils utils = new useExcel.ExcelUtils();

        utils.setExcelFile("C:\\Users\\Diana\\IdeaProjects\\group\\src\\test\\java\\testDate\\TestDate.xlsx", "first");

     int iTestCaseRow = utils.getRowContains("user0",0);
        System.out.println(iTestCaseRow);

      //  Object[][] testObjArray = utils.getTableArray( iTestCaseRow);

        return utils.getTableArray( iTestCaseRow);

    }
}