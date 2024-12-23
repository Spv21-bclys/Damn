import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelCombinationCounter {
    public static void main(String[] args) {
        String inputFilePath = "input.xlsx"; // Replace with your input file path
        String outputFilePath = "output.xlsx"; // Replace with your desired output file path

        // Column names to search for
        String column1Name = "Column1";
        String column2Name = "Column2";

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
            Map<String, Integer> combinationCountMap = new HashMap<>();

            // Get the header row (first row)
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("Header row is missing!");
                return;
            }

            // Find the column indexes based on the column names
            int column1Index = -1;
            int column2Index = -1;

            for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
                String headerValue = headerRow.getCell(i).toString().trim();
                if (headerValue.equalsIgnoreCase(column1Name)) {
                    column1Index = i;
                }
                if (headerValue.equalsIgnoreCase(column2Name)) {
                    column2Index = i;
                }
            }

            // If any of the columns are missing
            if (column1Index == -1 || column2Index == -1) {
                System.out.println("One or both columns are missing in the header.");
                return;
            }

            // Loop through rows and count combinations of Column1 and Column2
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Skip header row
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell1 = row.getCell(column1Index);
                Cell cell2 = row.getCell(column2Index);

                String value1 = cell1 != null ? cell1.toString().trim() : "";
                String value2 = cell2 != null ? cell2.toString().trim() : "";

                // Create combination string of Column1 and Column2
                String combination = value1 + "," + value2;
                // Count occurrences of each combination
                combinationCountMap.put(combination, combinationCountMap.getOrDefault(combination, 0) + 1);
            }

            // Write the results to a new Excel file
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Combinations");

            // Create header row for the output
            Row headerRowOutput = outputSheet.createRow(0);
            headerRowOutput.createCell(0).setCellValue("Combination");
            headerRowOutput.createCell(1).setCellValue("Count");

            // Populate data rows with combinations and counts
            int rowIndex = 1;
            for (Map.Entry<String, Integer> entry : combinationCountMap.entrySet()) {
                Row dataRow = outputSheet.createRow(rowIndex++);
                dataRow.createCell(0).setCellValue(entry.getKey());
                dataRow.createCell(1).setCellValue(entry.getValue());
            }

            // Save the output file
            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                outputWorkbook.write(fos);
            }

            System.out.println("Combination counts have been written to " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
