import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadData {
    public static void main(String[] args) {
        String excelFilePath = "Task-13.xlsx";
        try(FileInputStream fis = new FileInputStream(excelFilePath)) {

            //Create Workbook instance
            Workbook workbook = new HSSFWorkbook(fis);
            //Access the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            //iterate through the rows
            for(Row row : sheet) {
                for(Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING :
                            System.out.println(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.println((int) cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            break;
                    }
                }

                System.out.println();
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
