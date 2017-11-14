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
 * Created by Diana on 11.01.2017.
 */
public class ExcelUtils {
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;
    private static XSSFCell cell;
    private static Object[][] data;

    public static void setExcelFile(String fileName, String SheetName) throws IOException {
        try {
            FileInputStream ExcelFile = new FileInputStream(fileName);
            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);

        } catch (Exception e) {
            System.out.println("Exception" + e.getMessage());
        }
    }

    public static Object[][] getTableArray(int iTestCaseRow) throws Exception {
        String[][] tabArray = null;
        int startCol = 1;
        int ci=0,cj=0;
        int totalRows = 1;
        int totalCols = 2;

        tabArray=new String[totalRows][totalCols];

        for (int j=startCol;j<=totalCols;j++, cj++)
        {
            tabArray[ci][cj]=getCellData(iTestCaseRow,j);

            System.out.println(tabArray[ci][cj]);
        }
        return tabArray;
    }

    public static String getCellData(int rowIndex, int columnIndex) throws Exception {

        try {
            Cell = ExcelWSheet.getRow(rowIndex).getCell(columnIndex);
        } catch (Exception e) {

            System.err.println("The cell with row '" + rowIndex + "' and column '"
                    + columnIndex + "' doesn't exist in the sheet");

        }
        return new DataFormatter().formatCellValue(Cell);
    }
    public List getRowData() throws Exception {
        ArrayList<Object> dataList = new ArrayList<Object>();


        int i = 1;// excluding header row
        int totalRows = 6;
        while (i < totalRows) {
            System.out.println("loginToAppWithAllRoles : line : " + i);

            Object[] dataLine = new Object[2];
            dataLine[0] = getCellData(i, 0);
            dataLine[1] = getCellData(i, 1);

            dataList.add(dataLine);

            i++;
        }
        return  dataList;
    }


    public static int  getRowContains(String sTestCaseName, int colNum) throws Exception {
       int i;
        int rowCount = getRowUsed();
        for ( i = 0; i <= rowCount; i++) {
            String cellData = getCellData(i, colNum);
            if (cellData.equalsIgnoreCase(sTestCaseName)) {
                return i;
            }
        }
        return 0;
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


    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append("[");

        for(int i=0; i< data.length;i++){
            builder.append(data[i]+"\n");
            builder.append("]");

        }
        return builder.toString();

    }

    public static void main (String[]args) throws Exception {

     setExcelFile("C:\\Users\\Diana\\IdeaProjects\\group\\src\\test\\java\\testDate\\TestDate.xlsx", "first");

       // System.out.println(  getCellData(3,1));

           // System.out.println("Строка  с нужным названием" + "   " + getRowContains("user0", 0));

      int iTestCaseRow = getRowContains("user0",0);

        System.out.println("Название тест-метода соответствует строкам      "+ "  "+ iTestCaseRow);
        System.out.println("Колличество строк в файле"+ "   " +getRowUsed());


        System.out.println(   getTableArray(iTestCaseRow).toString() );



    }
}

     /*   XSSFSheet sheet =ExcelWBook.getSheet("first");
        XSSFRow row;
        XSSFCell cell;

        String result = "";

        Iterator rows = ExcelWSheet.rowIterator();
        while (rows.hasNext()) {
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (XSSFCell) cells.next();

                int cellType = cell.getCellType();
                switch (cellType) {
                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        break;
                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;

                    case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

        return result;
    }
*/

  /*  public static String getTestCaseName(String sTestCase)throws Exception{
      String value = sTestCase;
        try{
            int posi = value.indexOf("@");
           value = value.substring(0, posi);
            posi = value.lastIndexOf(".");
            value = value.substring(posi + 1);
            return value;
        }catch (Exception e){
            throw (e);
       }
    }
    */