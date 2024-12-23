import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelCombinationCounter {
    public static void main(String[] args) {
        String inputFilePath1 = "file1.xlsx"; // First file path
        String inputFilePath2 = "file2.xlsx"; // Second file path
        String outputFilePath = "output_comparison.xlsx"; // Output file path

        // Column names for both files
        String column1NameFile1 = "Column1"; // Column1 name in File1
        String column2NameFile1 = "Column2"; // Column2 name in File1

        String column1NameFile2 = "Col1"; // Column1 name in File2
        String column2NameFile2 = "Col2"; // Column2 name in File2

        // Count combinations in the first file
        Map<String, Integer> combinationCountFile1 = getCombinationCounts(inputFilePath1, column1NameFile1, column2NameFile1);

        // Count combinations in the second file
        Map<String, Integer> combinationCountFile2 = getCombinationCounts(inputFilePath2, column1NameFile2, column2NameFile2);

        // Create output workbook and sheet
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("Combination Comparison");

            // Create header row for the output
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("Combination");
            headerRow.createCell(1).setCellValue("Count (File1)");
            headerRow.createCell(2).setCellValue("Count (File2)");

            // Populate data rows with combinations and counts
            int rowIndex = 1;
            for (Map.Entry<String, Integer> entry : combinationCountFile1.entrySet()) {
                Row dataRow = outputSheet.createRow(rowIndex++);
                String combination = entry.getKey();
                Integer countFile1 = entry.getValue();
                Integer countFile2 = combinationCountFile2.getOrDefault(combination, 0);

                dataRow.createCell(0).setCellValue(combination);
                dataRow.createCell(1).setCellValue(countFile1);
                dataRow.createCell(2).setCellValue(countFile2);
            }

            // Save the output file
            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                outputWorkbook.write(fos);
            }

            System.out.println("Combination comparison has been written to " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Map<String, Integer> getCombinationCounts(String inputFilePath, String column1Name, String column2Name) {
        Map<String, Integer> combinationCountMap = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("Header row is missing!");
                return combinationCountMap;
            }

            // Resolve column indices from column names
            int column1Index = getColumnIndexByName(headerRow, column1Name);
            int column2Index = getColumnIndexByName(headerRow, column2Name);

            if (column1Index == -1 || column2Index == -1) {
                System.out.println("One or both column names not found!");
                return combinationCountMap;
            }

            // Loop through rows and count combinations of Column1 and Column2
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Skip header row
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell1 = row.getCell(column1Index);
                Cell cell2 = row.getCell(column2Index);

                String value1 = cell1 != null ? cell1.toString().trim() : "";
                String value2 = cell2 
