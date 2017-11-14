import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;


/**
 * Created by Diana on 08.12.2016.
 */
public class DatePrice extends FileNotFoundException {
    public static WebDriver driver;
    private static Logger Log = Logger.getLogger(DatePrice.class.getName());
    private static ArrayList<File> result = new ArrayList(); // ЛИСТ ДЛЯ ФАЙЛОВ



    @DataProvider(name = "Authentication")
    public Object[][] simpleDataProvider() {
        return new Object[][]{
                {"НАФТИЗИН"},
                {"АНАЛЬГИН ДЛЯ ДЕТЕЙ"},
                {"АНАЛЬГИН-ДАРНИЦА"},
                {"АНАЛЬГИН-ЗДОРОВЬЕ"}
        };
    }


    @Test(priority = 0)
    public static void contecst() throws ParseException {
        WebElement element = driver.findElement(By.xpath("//*[@id=\"navigation\"]/ul/li[3]/a"));

        Assert.assertTrue(element.isDisplayed());
        element.click();
        driver.navigate().forward();
        String actualUrl = "http://pharmbase.com.ua/poisk1/";
        String currentUrl = driver.getCurrentUrl();

        Assert.assertTrue(actualUrl.contentEquals(currentUrl));

        if (actualUrl.equals(currentUrl)) {
            System.out.println("Verification Successful - The correct Url is opened.");
            Log.info("*******************************************************************************");

            Reporter.log("Verification Successful - The correct Url is opened | ");
        } else {
            System.out.println("Actual URL is : " + actualUrl);
            System.out.println("Expected URL is : " + currentUrl);
            Log.info("*******************************************************************************");

            Reporter.log("Verification Failed - The correct Url is not open | ");
        }

        WebElement element1 = driver.findElement(By.name("N"));
        element1.click();
        element1.sendKeys("Анальгин");

        List<WebElement> rdButton = driver.findElements(By.name("PM"));

        boolean bValue = false;
        // This statement will return True, in case of first Radio button is selected
        bValue = rdButton.get(1).isSelected();

        // This will check that if the bValue is True means if the first radio button is selected
        if (bValue == true) {
            // This will select Second radio button, if the first radio button is selected by default
            rdButton.get(1).click();
        } else {
            // If the first radio button is not selected by default, the first will be selected
            rdButton.get(0).click();
        }
        WebElement element2 = driver.findElement(By.xpath("//*[@id=\"content\"]/form[1]/table/tbody/tr[4]/td[2]/input"));
        element2.submit();
        driver.navigate().forward();
        String resultUrl = "http://pharmbase.com.ua/result/?N=%D0%90%D0%BD%D0%B0%D0%BB%D1%8C%D0%B3%D0%B8%D0%BD&PM=1&PS=500";

        String currentResultUrl = driver.getCurrentUrl();
        if (resultUrl.equals(currentResultUrl)) {
            Log.info("*******************************************************************************");
            System.out.println("Verification Successful - The correct Url is opened.");
            Reporter.log("Verification Successful - The correct Url is opened | ");
        } else {
            Log.info("*******************************************************************************");
            System.out.println("Actual URL is : " + currentResultUrl);
            System.out.println("Expected URL is : " + resultUrl);

            Reporter.log("Verification Failed - The correct Url is not open | ");
        }

        // Содержание страницы
        WebElement element15 = driver.findElement(By.xpath("//*[@id=\"main\"]"));

        Assert.assertTrue(element15.isDisplayed());
        Log.info("*******************************************************************************");

        //  System.out.println(element15.getText());

        Reporter.log("Verification Successul - The correct info is displayed | ");

    }

    // ПРОВЕРКА ДАТЫ ПРАЙСА

    @Test(priority = 1)
    public void check_data() {
        LocalDate toDate = LocalDate.now();
        String val;
        String val_submit;
        WebElement value_date1 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[2]/td[1]"));
        WebElement value_date2 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[3]/td[1]"));
        WebElement value_date3 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[4]/td[1]"));
        WebElement value_date4 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[5]/td[1]"));
        WebElement value_date5 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[6]/td[1]"));
        WebElement value_date6 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[7]/td[1]"));
        WebElement value_date7 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[8]/td[1]"));
        WebElement value_date8 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[9]/td[1]"));
        WebElement value_date9 = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[10]/td[1]"));

        ArrayList<WebElement> list_data = new ArrayList<WebElement>();
        list_data.add(value_date1);
        list_data.add(value_date2);
        list_data.add(value_date3);
        list_data.add(value_date4);
        list_data.add(value_date5);
        list_data.add(value_date6);
        list_data.add(value_date7);
        list_data.add(value_date8);
        list_data.add(value_date9);

        for (int i = 0; i < list_data.size(); i++) {
            val = list_data.get(i).getText();
            val = val.substring(0, 10);
            //System.out.println(val_submit);
        }

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        val_submit = toDate.format(formatter);
        LocalDate dateLocal = LocalDate.parse(val_submit, formatter);


        Period period = Period.between(dateLocal, toDate);

        for (int j = 0; j < list_data.size(); j++) {
            if (period.getDays() > 14) {
                Log.info("*******************************************************************************");
                System.out.println("Предложения старше двух недель от текущей даты " + " " + period);

                Reporter.log("Verification Failed - Предложения старше двух недель отображаются | ");

            }
        }
        Log.info("*******************************************************************************");
        System.out.println("Предложения старше двух недель отсутствуют");
        Reporter.log("Verification Successul- Предложения старше двух недель отсутствуют| ");

    }
    // ЗАГРУЗКА ПРАЙСА

    @Test(priority = 3)
    public void findFile( ) throws InterruptedException {
        File dir= new File("C:\\Users\\Diana\\Downloads"); // путь к директории
        WebElement price = driver.findElement(By.xpath("//*[@id=\"content\"]/table[3]/tbody/tr[2]/td[2]/a"));
        price.click();
        Thread.sleep(1000);
        if (!dir.exists()) {
            System.out.println("Неверный путь, попробуйте еще.");
                        return;
                   } System.out.println("Файл для поиска: " + price + "\n" +
                                "Каталог в котором ищем: " + dir);
                find(dir);

                if (result.isEmpty()) {
                        System.out.println("По указаному пути файл не найден.");
                    } else {
                        System.out.println("Файл найден по следующему пути:");
                        for (File f: result) {
                                System.out.println(f.getPath());
                            }
                        System.out.println("Всего найдено файлов: " + result.size());
                    }

                    }

  //ПОИСК ФАЙЛА В КАТАЛОГЕ
    private static void find(File dir) {
        if (dir.exists() && dir.isDirectory()) {
            if (dir.listFiles().length > 0) {
                for (File f : dir.listFiles()) {
                    if (f.getName().indexOf('.') != -1 || f.getName().substring(f.getName().lastIndexOf('.') + 1).equals("xlsx")) {

                        result.add(f);
                        System.out.println(f.getName());


                    } else {

                       System.out.println("Неверный формат файла.");
                        return ;
                    }

                }

            }

        }

    }

    // ПРОВЕРКА КОТЕНТА СТРАНИЦЫ + ПОИСК НУЖНОГО СТОЛБЦА В ТАБЛИЦЕ И ВЫВОД ДАННЫХ
    @Test(priority = 4)
    public void check_rbName() {
        Log.info("Драйвер вернулся на страницу поиска");
        driver.navigate().to("http://pharmbase.com.ua/poisk1/");
        Log.info("Драйвер выбрал чек поинт Поиск с синонимами");
        List<WebElement> rbList = driver.findElements(By.name("RT"));
        boolean value = false;
        value = rbList.get(0).isSelected();
        if (value = true) {
            rbList.get(0).click();

            // driver.findElement(By.xpath("//*[@id=\"content\"]/form[1]/table/tbody/tr[4]/td[2]/input")).click();
            WebElement element1 = driver.findElement(By.name("N"));
            Log.info("Драйвер нашел форму поиска препарата, ввел название препарата");
            element1.click();
            element1.clear();
            element1.sendKeys("Анальгин");
            WebElement element2 = driver.findElement(By.xpath("//*[@id=\"content\"]/form[1]/table/tbody/tr[1]/td[2]/input"));
            element2.click();
            Log.info("Драйвер перешел на страницу поиска синонимов по введенному препарату");

            String url = "http://pharmbase.com.ua/result/?N=%D0%B0%D0%BD%D0%B0%D0%BB%D1%8C%D0%B3%D0%B8%D0%BD&PM=1&RT=6&PS=500";
            driver.get(url);
            Assert.assertTrue(url.equals("http://pharmbase.com.ua/result/?N=%D0%B0%D0%BD%D0%B0%D0%BB%D1%8C%D0%B3%D0%B8%D0%BD&PM=1&RT=6&PS=500"));
            Log.info("Проверка контекста страницы");
            if (url.contentEquals("http://pharmbase.com.ua/result/?N=%D0%B0%D0%BD%D0%B0%D0%BB%D1%8C%D0%B3%D0%B8%D0%BD&PM=1&RT=6&PS=500")) {
                System.out.println("Контент страницы соответствует запрашиваемому");
            } else {
                System.out.println("Контент не соответствует нужной странице");
            }


            Log.info("Поиск таблицы с препаратами - синонимами");
            WebElement table = driver.findElement(By.xpath("//*[@id=\"content\"]/table[2]"));

          // ПОИСК ЗНАЧЕНИЙ ИЗ СТОЛБЦА "НАИМЕНОВАНИЯ"
            Log.info("Получить значения  в ячейках таблицы");
            Log.info("Поиск столбца Названия, вывод данных");
            for (int i = 2; i <= 6; i++) {
                String column_value = null;
                column_value = driver.findElement(By.xpath("//*[@id=\"content\"]/table[2]/tbody/tr[" + i + "]/td[2]")).getText();
                if (column_value.equalsIgnoreCase("Наименование")) {
                    System.out.println(column_value);
                }
                //  ПОЛУЧЕНИЕ ЗНАЧЕНИЙ ИЗ СТОЛБЦА "НАИМЕНОВАНИЯ"

            for (int j = 1; j <= 2; j++) {
                String cell_value = driver.findElement(By.xpath("//*[@id=\"content\"]/table[2]/tbody/tr[ "+ i +"]/td[" + j + "]")).getText();
                System.out.println(cell_value);
            }

        }

    }
}
       //ПРОВЕРКА НА СООТВЕТСТВИЕ СИНОНИМОВ
     @Test(dataProvider = "Authentication", priority = 5)
     public void equels_date(String simplParame) {
         //String cell_value = "";
         for (int i = 3; i <= 6; i++) {
             String  cell_value = driver.findElement(By.xpath("//*[@id=\"content\"]/table[2]/tbody/tr[" + i + "]/td[2]")).getText();
             for (int j = 0; j <= simpleDataProvider().length; j++) {

                 if (simplParame.equalsIgnoreCase(cell_value)) {
                     System.out.println("Синоним значения" + "  " + cell_value + "  " + "Соотвтствует значению" + "  " + simplParame);
                       return;
                 }else{
                     System.out.println(simpleDataProvider().toString()+"Значение не корректно" +" "+ cell_value);
                     return;

                 }
             }

         }

     }

  @BeforeTest
    public void beforeMethod() {
        DOMConfigurator.configure("log4j.xml");

        driver = new ChromeDriver();
        Log.info("New driver instantiated");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        driver.get("http://pharmbase.com.ua/");
        Log.info("Web application launched");

    }
    @AfterTest
    public void afterTest(){
        driver.close();
    }

}
//*[@id="content"]/table/tbody/tr[1]/td[2]