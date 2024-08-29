import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteData {
        public static void main(String[] args) {
            //create a workbook new
            Workbook workbook = new HSSFWorkbook();

            //create a sheet new called "Sheet1"
            Sheet sheet = workbook.createSheet("Sheet1");

            //Create a Header Row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("Email");

            //Write data to rows
            Object[][] data = {
                    {"John Doe", 30, "john@test.com"},
                    {"Jane Doe", 28, "john@test.com"},
                    {"Bob Smith", 35, "jacky@example.com"},
                    {"Swapnil", 37, "swapnil@example.com"}
            };

            int rowCount = 1;
            for(Object[] rowData : data) {
                Row row = sheet.createRow(rowCount++);
                int columnCount = 0;
                for(Object field : rowData) {
                    Cell cell = row.createCell(columnCount++);
                    if(field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);

                    }
                }
            }

            //Write the output to a file
            try (FileOutputStream outputStream = new FileOutputStream("Task-13.xlsx")) {
                workbook.write(outputStream);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            System.out.println("Data Written to Excel File Successfully");

        }

}
