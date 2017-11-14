package poi;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

/**
 * Created by Diana on 30.12.2016.
 */
public class ReadExcelDataProvider {
    public WebDriver driver;
    public WebDriverWait wait;
    String appUrl = "http://pharmbase.com.ua/";

    //Locators
    private By byEmail = By.name("login");
    private By byPassword = By.name("pass");
    private By bySubmit = By.xpath("//*[@id=\"vxod-na-sajt\"]/form/table/tbody/tr[3]/td[2]/input");
    private By byError = By.xpath("");

    @BeforeClass
    public void setSetup() {
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, 5);
        driver.get(appUrl );
    }

    @Test(dataProvider = "empLogin")
    public void VerifyInvalidLogin(String userName,String password) {


        driver.findElement(byEmail).sendKeys(userName);
        driver.findElement(byPassword).sendKeys(password);

        //wait for element to be visible and perform click
    //  wait.until(ExpectedConditions.visibilityOfElementLocated(bySubmit));

        driver.findElement(bySubmit).click();
        //Check for error message
     // wait.until(ExpectedConditions.visibilityOfElementLocated(byError))


      String actualErrorDisplayed = driver.findElement(byError).getText();
      String requiredErrorMessage = "Please correct the marked field(s) below.";

      Assert.assertEquals(requiredErrorMessage, actualErrorDisplayed);


    }

    @DataProvider(name = "empLogin")
    public Object[][] loginData() throws FileNotFoundException, BiffException {
        Object[][] arrayObject = getExcelData("C:\\Users\\Diana\\IdeaProjects\\group\\src\\test\\java\\testDate\\Test.xls", "Лист2");
        return arrayObject;
    }

    public String[][] getExcelData(String fileName, String sheetName) throws FileNotFoundException, BiffException {
        String[][] arrayExcelData = null;
        try {
            FileInputStream fs = new FileInputStream(fileName);
            Workbook wb = Workbook.getWorkbook(fs);
            Sheet sh = wb.getSheet(sheetName);
            int totalNoOfCols = sh.getColumns();
            int totalNoOfRows = sh.getRows();

            arrayExcelData = new String[totalNoOfRows - 1][totalNoOfCols];
            for (int i = 1; i < totalNoOfRows; i++) {

                for (int j = 0; j < totalNoOfCols; j++) {
                    arrayExcelData[i - 1][j] = sh.getCell(j, i).getContents();
                }

            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
        return arrayExcelData;
    }
    @AfterClass
public void afterTest()
    {
    driver.quit();
}

}
