package File;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Diana on 30.01.2017.
 */
public class CreateEx {
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;
    int rowNum ;
    List <DataModel> dataList = new ArrayList();


    private List<DataModel> fillData() {
        List<DataModel> dataModels = new ArrayList<>();


        dataModels.add(new DataModel("Howard", "Wolowitz", "Massachusetts", 90000.0));
        dataModels.add(new DataModel("Leonard", "Hofstadter", "Massachusetts", 95000.0));
        dataModels.add(new DataModel("Sheldon", "Cooper", "Massachusetts", 120000.0));
      //  return dataModels;

        for (int i = 0; i < dataModels.size(); i++) {
            dataList.add(dataModels.get(i));

        }
        return dataList;// аполняем лист данными
    }


    private void  createSheetHeader(HSSFSheet sheet, int rowNum, DataModel dataModel){

        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(dataModel.getName());
        row.createCell(1).setCellValue(dataModel.getSurname());
        row.createCell(2).setCellValue(dataModel.getCity());
        row.createCell(3).setCellValue(dataModel.getSalary());
        }

    private void setExel(String path) throws IOException {
        try {

            FileOutputStream out = new FileOutputStream(new File(path));
            workbook = new HSSFWorkbook();
            sheet = workbook.createSheet("list1");
            rowNum=0;
            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue("Имя");
            row.createCell(1).setCellValue("Фамилия");
            row.createCell(2).setCellValue("Город");
            row.createCell(3).setCellValue("Зарплата");

            fillData();

            for (DataModel dataModel : dataList) {
                createSheetHeader(sheet, ++rowNum, dataModel);
            }

            workbook.write(out);


        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }

    public static void main(String[] args) throws IOException {
        CreateEx createEx = new CreateEx();

        createEx.setExel("D:\\Program\\IdeaProjects\\group\\src\\main\\java\\File.xls");

    }
    }

