package table.Project5;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelFileGenerator {

    public static void main(String[] args) {
        Date fromDate = new Date(); // Replace with your from date
        Date toDate = new Date(); // Replace with your to date

        // Format the dates to include in the filename
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        String fromDateStr = dateFormat.format(fromDate);
        String toDateStr = dateFormat.format(toDate);

        String fileName = "Report_" + fromDateStr.replace("-", "/") + "_to_" + toDateStr.replace("-", "/") + ".xlsx";

        // Create Excel workbook and sheet
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Example: Write some data to the Excel file
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Column 1");
            headerRow.createCell(1).setCellValue("Column 2");

            // Example: Write data rows
            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("Data 1");
            dataRow.createCell(1).setCellValue("Data 2");

            // Write the workbook content to a file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
                System.out.println("Excel file created successfully with filename: " + fileName);
            } catch (IOException e) {
                System.out.println("Error writing Excel file: " + e.getMessage());
            }
        } catch (IOException e) {
            System.out.println("Error creating Excel workbook: " + e.getMessage());
        }
    }
}
