import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelValidator {
    public static void main(String[] args) {
        String ruleFilePath = "path_to_rule_book.xlsx";
        String inputFilePath = "path_to_input_excel.xlsx";
        String outputFilePath = "output_excel.xlsx";

        try {
            Map<String, List<Map<String, String>>> rules = readRules(ruleFilePath);
            validateAndGenerateOutput(inputFilePath, outputFilePath, rules);
            System.out.println("Validation completed. Results saved to: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

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

    private static void validateAndGenerateOutput(String inputFilePath, String outputFilePath,
                                                  Map<String, List<Map<String, String>>> rules) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = workbook.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet("Validated Output");

            // Copy header
            Row headerRow = inputSheet.getRow(0);
            Row outputHeaderRow = outputSheet.createRow(0);
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                outputHeaderRow.createCell(i).setCellValue(headerRow.getCell(i).getStringCellValue());
            }

            int outputRowNum = 1;

            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row inputRow = inputSheet.getRow(i);
                if (inputRow == null) continue;

                String billingCode = getCellValue(inputRow.getCell(findColumnIndex(headerRow, "BIILING_CODE")));
                List<Map<String, String>> ruleSets = rules.getOrDefault(billingCode, new ArrayList<>());

                boolean ruleMatched = false;

                for (Map<String, String> ruleSet : ruleSets) {
                    boolean isValid = validateRowAgainstRule(inputRow, headerRow, ruleSet);
                    if (isValid) {
                        appendRowToSheet(outputSheet, inputRow, outputRowNum++);
                        ruleMatched = true;
                        break; // Only one rule match should be appended
                    }
                }

                if (!ruleMatched && !ruleSets.isEmpty()) {
                    // If no rules match, append the first rule arbitrarily
                    appendRowToSheet(outputSheet, inputRow, outputRowNum++);
                }
            }

            // Save the updated Excel file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
            }
        }
    }

    private static boolean validateRowAgainstRule(Row row, Row headerRow, Map<String, String> ruleSet) {
        for (Map.Entry<String, String> rule : ruleSet.entrySet()) {
            String columnName = rule.getKey();
            String expectedValue = rule.getValue();
            int colIndex = findColumnIndex(headerRow, columnName);
            String actualValue = getCellValue(row.getCell(colIndex));

            if (!validateCellValue(expectedValue, actualValue)) {
                return false;
            }
        }
        return true;
    }

    private static boolean validateCellValue(String expectedValue, String actualValue) {
        if (expectedValue.equalsIgnoreCase("Not Used")) return true;
        if (expectedValue.startsWith("<>")) {
            String[] excludedValues = expectedValue.substring(2).split(",");
            return Arrays.stream(excludedValues).noneMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        if (expectedValue.contains(",")) {
            String[] allowedValues = expectedValue.split(",");
            return Arrays.stream(allowedValues).anyMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        return expectedValue.equalsIgnoreCase(actualValue);
    }

    private static void appendRowToSheet(Sheet sheet, Row inputRow, int outputRowNum) {
        Row outputRow = sheet.createRow(outputRowNum);
        for (int i = 0; i < inputRow.getLastCellNum(); i++) {
            Cell inputCell = inputRow.getCell(i);
            Cell outputCell = outputRow.createCell(i);

            if (inputCell != null) {
                outputCell.setCellValue(getCellValue(inputCell));
            }
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }

    private static int findColumnIndex(Row headerRow, String columnName) {
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column " + columnName + " not found");
    }
}
