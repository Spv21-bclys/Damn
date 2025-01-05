import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelCombineAndCountThreeFilesInMemory {

    public static ByteArrayOutputStream processExcelFiles(ByteArrayInputStream inputStream1, 
                                                           ByteArrayInputStream inputStream2, 
                                                           ByteArrayInputStream inputStream3) throws IOException {
        // Read workbooks from ByteArrayInputStream
        try (Workbook workbook1 = new XSSFWorkbook(inputStream1);
             Workbook workbook2 = new XSSFWorkbook(inputStream2);
             Workbook workbook3 = new XSSFWorkbook(inputStream3)) {

            // Reading sheets from all three files
            Sheet sheet1 = workbook1.getSheetAt(0); // File 1 (Your Reference, Currency)
            Sheet sheet2 = workbook2.getSheetAt(0); // File 2 (Your Reference, Currency, Result)
            Sheet sheet3 = workbook3.getSheetAt(0); // File 3 (Your Reference, Currency, Result, Status)

            // Maps to store frequencies of combinations from Excel 1, Excel 2, and Excel 3
            Map<String, Integer> combinationCountsFromExcel1 = new HashMap<>();
            Map<String, Map<String, Integer>> combinationCountsFromExcel2 = new HashMap<>();
            Map<String, Map<String, Integer>> combinationCountsFromExcel3 = new HashMap<>();

            // Sets to store unique Result and Status values
            Set<String> resultValues = new HashSet<>();
            Set<String> statusValues = new HashSet<>();

            // Step 1: Get header row and find column indices for column names
            Row headerRow1 = sheet1.getRow(0);
            Row headerRow2 = sheet2.getRow(0);
            Row headerRow3 = sheet3.getRow(0);

            int refIndex1 = getColumnIndex(headerRow1, "Your Reference");
            int currencyIndex1 = getColumnIndex(headerRow1, "Currency");

            int refIndex2 = getColumnIndex(headerRow2, "Your Reference");
            int currencyIndex2 = getColumnIndex(headerRow2, "Currency");
            int resultIndex2 = getColumnIndex(headerRow2, "Result");

            int refIndex3 = getColumnIndex(headerRow3, "Your Reference");
            int currencyIndex3 = getColumnIndex(headerRow3, "Currency");
            int resultIndex3 = getColumnIndex(headerRow3, "Result");
            int statusIndex3 = getColumnIndex(headerRow3, "Status");

            // Step 2: Count frequencies of combinations in Excel 1
            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row = sheet1.getRow(i);
                if (row != null) {
                    String reference = row.getCell(refIndex1).getStringCellValue();
                    String currency = row.getCell(currencyIndex1).getStringCellValue();

                    // Create a unique key for the combination
                    String combination = reference + "-" + currency;

                    // Update the count in the map
                    combinationCountsFromExcel1.put(combination,
                            combinationCountsFromExcel1.getOrDefault(combination, 0) + 1);
                }
            }

            // Step 3: Count frequencies of combinations for each Result value in Excel 2
            for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
                Row row = sheet2.getRow(i);
                if (row != null) {
                    String reference = row.getCell(refIndex2).getStringCellValue();
                    String currency = row.getCell(currencyIndex2).getStringCellValue();
                    String resultValue = row.getCell(resultIndex2).getStringCellValue();

                    // Create a unique key for the combination
                    String combination = reference + "-" + currency;

                    // Ensure the result value is tracked
                    resultValues.add(resultValue);

                    // Initialize the combination map if it doesn't exist
                    combinationCountsFromExcel2.putIfAbsent(combination, new HashMap<>());

                    // Update the count for the specific result value
                    Map<String, Integer> resultMap = combinationCountsFromExcel2.get(combination);
                    resultMap.put(resultValue, resultMap.getOrDefault(resultValue, 0) + 1);
                }
            }

            // Step 4: Count frequencies of combinations for each Status value in Excel 3
            for (int i = 1; i <= sheet3.getLastRowNum(); i++) {
                Row row = sheet3.getRow(i);
                if (row != null) {
                    String reference = row.getCell(refIndex3).getStringCellValue();
                    String currency = row.getCell(currencyIndex3).getStringCellValue();
                    String resultValue = row.getCell(resultIndex3).getStringCellValue();
                    String statusValue = row.getCell(statusIndex3).getStringCellValue();

                    // Create a unique key for the combination
                    String combination = reference + "-" + currency;

                    // Ensure the status value is tracked
                    statusValues.add(statusValue);

                    // Initialize the combination map if it doesn't exist
                    combinationCountsFromExcel3.putIfAbsent(combination, new HashMap<>());

                    // Update the count for the specific status value
                    Map<String, Integer> statusMap = combinationCountsFromExcel3.get(combination);
                    statusMap.put(statusValue, statusMap.getOrDefault(statusValue, 0) + 1);
                }
            }

            // Step 5: Write the results to a new Excel file (in memory)
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Unique Combinations");

            // Header row
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("Combination");
            headerRow.createCell(1).setCellValue("Count");

            // Write headers for Result values
            List<String> resultValueList = new ArrayList<>(resultValues);
            for (int i = 0; i < resultValueList.size(); i++) {
                headerRow.createCell(i + 2).setCellValue("Frequency for " + resultValueList.get(i));
            }

            // Write headers for Status values
            List<String> statusValueList = new ArrayList<>(statusValues);
            for (int i = 0; i < statusValueList.size(); i++) {
                headerRow.createCell(i + 2 + resultValueList.size()).setCellValue("Frequency for " + statusValueList.get(i));
            }

            // Populate the result sheet
            int rowNum = 1;
            for (String combination : combinationCountsFromExcel1.keySet()) {
                Row row = outputSheet.createRow(rowNum++);
                row.createCell(0).setCellValue(combination);

                // Frequency from Excel 1
                int count = combinationCountsFromExcel1.getOrDefault(combination, 0);
                row.createCell(1).setCellValue(count);

                // Frequencies for each Result value from Excel 2
                Map<String, Integer> resultMap = combinationCountsFromExcel2.getOrDefault(combination, new HashMap<>());
                for (int i = 0; i < resultValueList.size(); i++) {
                    String resultValue = resultValueList.get(i);
                    int frequency = resultMap.getOrDefault(resultValue, 0);
                    row.createCell(i + 2).setCellValue(frequency);
                }

                // Frequencies for each Status value from Excel 3
                Map<String, Integer> statusMap = combinationCountsFromExcel3.getOrDefault(combination, new HashMap<>());
                for (int i = 0; i < statusValueList.size(); i++) {
                    String statusValue = statusValueList.get(i);
                    int frequency = statusMap.getOrDefault(statusValue, 0);
                    row.createCell(i + 2 + resultValueList.size()).setCellValue(frequency);
                }
            }

            // Write to ByteArrayOutputStream
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            outputWorkbook.write(byteArrayOutputStream);
            return byteArrayOutputStream;
        }
    }

    // Helper method to get column index by name
    private static int getColumnIndex(Row headerRow, String columnName) {
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column name " + columnName + " not found in header.");
    }

    // Example usage:
    public static void main(String[] args) throws IOException {
        // Replace these with actual ByteArrayOutputStream of your Excel files
        ByteArrayOutputStream outputStream1 = new ByteArrayOutputStream(); // File 1
        ByteArrayOutputStream outputStream2 = new ByteArrayOutputStream(); // File 2
        ByteArrayOutputStream outputStream3 = new ByteArrayOutputStream(); // File 3

        // Convert ByteArrayOutputStream to ByteArrayInputStream for processing
        ByteArrayInputStream inputStream1 = new ByteArrayInputStream(outputStream1.toByteArray());
        ByteArrayInputStream inputStream2 = new ByteArrayInputStream(outputStream2.toByteArray());
        ByteArrayInputStream inputStream3 = new ByteArrayInputStream(outputStream3.toByteArray());

        // Process the Excel files
        ByteArrayOutputStream resultStream = processExcelFiles(inputStream1, inputStream2, inputStream3);

        // The result is now in 'resultStream', which can be saved to a file or used as needed.
        try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
            resultStream.writeTo(fos);
        }
    }
}
