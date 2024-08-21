package JavaQA.ExcelCreation;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadDataFromExcel {

    public static void main(String[] args) throws IOException {

        // Create a new workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a new sheet
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        // Create header row
        XSSFRow headerRow = sheet.createRow(0);

        // Write header cells
        String[] headers = {"Name", "Age", "Email"};
        for (int i = 0; i < headers.length; i++) {
            XSSFCell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
        }

        // Create data rows
        String[][] data = {
                {"John Doe", "30", "john@test.com"},
                {"Jane Doe", "28", "jane@test.com"},
                {"Bob Smith", "35", "jacky@example.com"},
                {"Swapnil", "37", "swapnil@example.com"}
        };

        for (int i = 0; i < data.length; i++) {
            XSSFRow dataRow = sheet.createRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                XSSFCell dataCell = dataRow.createCell(j);
                dataCell.setCellValue(data[i][j]);
            }
        }

        // Write the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream("data.xlsx")) {
            workbook.write(outputStream);
        }

        System.out.println("Data written to Excel file successfully!");
    }
}