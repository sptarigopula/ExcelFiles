import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * Created by starigopula on 3/8/2018.
 */
public class WritingToExcelFile {

    public static void main(String args[]) throws IOException {
        try {
            String excelFilePath = "C://Users//starigopula//ReadingFromExcel.xlsx";
            File file = new File(excelFilePath);

            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(inputStream);

            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            Cell column = row.getCell(1);

            String updatename="Lalalala";
            column.setCellValue(updatename);

            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
