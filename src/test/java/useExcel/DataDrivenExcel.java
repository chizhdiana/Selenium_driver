package useExcel;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Diana on 19.01.2017.
 */
public class DataDrivenExcel {
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;
    private static XSSFCell cell;

    public static void setExcelFile(String fileName, String SheetName) throws IOException {
        try {
            FileInputStream ExcelFile = new FileInputStream(fileName);
            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);

        } catch (Exception e) {
            System.out.println("Exception" + e.getMessage());
        }
    }
    public static String getCellData(int row, int col) throws Exception {

        try {
            Cell = ExcelWSheet.getRow(row).getCell(col);
        } catch (Exception e) {

            System.err.println("The cell with row '" + row + "' and column '"
                    + col + "' doesn't exist in the sheet");

        }
       // return Cell.getStringCellValue();
        return new DataFormatter().formatCellValue(Cell);
    }

    public static int getRowUsed() throws Exception {

        try{
            int RowCount = ExcelWSheet.getLastRowNum();
            return RowCount;

        }catch (Exception e){
            System.out.println(e.getMessage());
            throw (e);
        }
    }

    public List getRowContains(String testCaseName, int colNum) throws Exception { // получим, по какому юзеру получать данніе
        List list = new ArrayList(); // параметры: название юзера и колонка в которой он находиться
        int rowCount = getRowUsed();
        for (int i = 0; i <= rowCount; i++) {
           String cellData = getCellData(i, colNum);
            if (cellData.equalsIgnoreCase(testCaseName)) {
                list.add(i);
            }
        }
        return list;
    }

    public List[] getRowData(int rowNo) throws Exception { // получить данные из указанной строки
        List[] arr = new List[1]; // один елемент в которій передаем лист
        List list = new ArrayList();

        int startCol = 1;
        int totalCols = ExcelWSheet.getRow(rowNo)// получаем нужную строку
                .getPhysicalNumberOfCells();;
             for (int i = startCol; i < totalCols; i++) {
                 String cellData = getCellData(rowNo, i); // rowNo - строка, i- перебирает все  ячейки в заданной строке
                 list.add(cellData);// добавили полученнные данные в лист и дальше ложим лист в аррай
           }
           arr[0]=list;
           return arr;

        }

    public Object[][] getTableArray(List dataList ) throws Exception {// лист передает данные по строке
     // ПРИМЕР   List dataList = userData.getRowContains("user0",0); // получим данные какой тесткейснайм
        Object[][] data = new Object[dataList.size()][];
        for (int i = 0; i < dataList.size(); i++) {
            data[i] = getRowData((Integer) dataList.get(i)); // данные из строки по заданному тесткейснейм
        }

            return data;
    }





    public static void main(String[] args) throws Exception {
        DataDrivenExcel userData = new DataDrivenExcel();
        setExcelFile("C:\\Users\\Diana\\IdeaProjects\\group\\src\\test\\java\\testDate\\TestDate.xlsx", "first");

      List dataList = userData.getRowContains("user0",0);
        System.out.println(dataList.toString());
        System.out.println(userData.getCellData(1,2));// параметры: первая строка, 2-я ячейкка
        System.out.println("Название тест-метода соответствует строкам      "+ "  "+  dataList);// массив 2 и 3 строка
        for(int i=0; i< dataList.size();i++) {
            System.out.println(userData.getTableArray(dataList));
        }

    }
}
