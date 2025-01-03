import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelValidator {
    public static void main(String[] args) {
        String ruleFilePath = "path_to_rule_book.xlsx";
        String inputFilePath = "path_to_excel_b.xlsx";
        String outputFilePath = "output_excel_b.xlsx";

        try {
            Map<String, List<Map<String, String>>> rules = readRules(ruleFilePath);
            validateAndUpdateExcelB(inputFilePath, outputFilePath, rules);
            System.out.println("Validation completed. Results saved to: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Read Rule Book and store as Map<String, List<Map<String, String>>>
    private static Map<String, List<Map<String, String>>> readRules(String filePath) throws IOException {
        Map<String, List<Map<String, String>>> rules = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            List<String> columnNames = new ArrayList<>();
            for (Cell cell : headerRow) {
                columnNames.add(cell.getStringCellValue());
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String billingCode = getCellValue(row.getCell(findColumnIndex(headerRow, "BIILING_CODE")));
                Map<String, String> ruleSet = new HashMap<>();

                for (String columnName : columnNames) {
                    int colIndex = findColumnIndex(headerRow, columnName);
                    String cellValue = getCellValue(row.getCell(colIndex));
                    ruleSet.put(columnName, cellValue);
                }

                rules.computeIfAbsent(billingCode, k -> new ArrayList<>()).add(ruleSet);
            }
        }
        return rules;
    }

    // Validate Excel B and update with results
    private static void validateAndUpdateExcelB(String inputFilePath, String outputFilePath,
                                                Map<String, List<Map<String, String>>> rules) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int validationColIndex = headerRow.getLastCellNum();

            // Add "Validation Result" column
            headerRow.createCell(validationColIndex).setCellValue("Validation Result");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String billingCode = getCellValue(row.getCell(findColumnIndex(headerRow, "BIILING_CODE")));
                List<Map<String, String>> ruleSets = rules.getOrDefault(billingCode, new ArrayList<>());

                boolean isValid = validateRowAgainstRules(row, headerRow, ruleSets);
                row.createCell(validationColIndex).setCellValue(isValid ? "Correct" : "Wrong");
            }

            // Save the updated Excel file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        }
    }

    // Validate a single row against multiple rule sets
    private static boolean validateRowAgainstRules(Row row, Row headerRow, List<Map<String, String>> ruleSets) {
        for (Map<String, String> ruleSet : ruleSets) {
            boolean isMatch = true;
            for (Map.Entry<String, String> rule : ruleSet.entrySet()) {
                String columnName = rule.getKey();
                String expectedValue = rule.getValue();
                int colIndex = findColumnIndex(headerRow, columnName);
                String actualValue = getCellValue(row.getCell(colIndex));

                if (!validateCellValue(expectedValue, actualValue)) {
                    isMatch = false;
                    break;
                }
            }
            if (isMatch) return true; // If one rule set matches, row is valid
        }
        return false; // All rule sets failed
    }

    // Validate cell value based on rule
    private static boolean validateCellValue(String expectedValue, String actualValue) {
        if (expectedValue.equalsIgnoreCase("Not Used")) return true; // Any value is valid
        if (expectedValue.startsWith("<>")) { // Exclusion rule
            String[] excludedValues = expectedValue.substring(3, expectedValue.length() - 1).split(",");
            return Arrays.stream(excludedValues).noneMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        if (expectedValue.contains(",")) { // Multiple values allowed
            String[] allowedValues = expectedValue.split(",");
            return Arrays.stream(allowedValues).anyMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        return expectedValue.equalsIgnoreCase(actualValue); // Exact match
    }

    // Helper method to get cell value as String
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }

    // Find column index by column name
    private static int findColumnIndex(Row headerRow, String columnName) {
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column " + columnName + " not found");
    }
}
