import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelValidator {

    public static ByteArrayOutputStream validateExcel(
            String ruleFilePath,
            ByteArrayOutputStream inputFileStream) throws IOException {

        // Read the rules from the rule file
        Map<String, List<Map<String, String>>> rules = readRules(ruleFilePath);

        // Validate and update the input Excel using the rules
        return validateAndUpdateExcel(new ByteArrayInputStream(inputFileStream.toByteArray()), rules);
    }

    private static Map<String, List<Map<String, String>>> readRules(String ruleFilePath) throws IOException {
        Map<String, List<Map<String, String>>> rules = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(ruleFilePath);
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

    private static ByteArrayOutputStream validateAndUpdateExcel(
            InputStream inputFileStream,
            Map<String, List<Map<String, String>>> rules) throws IOException {

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try (Workbook inputWorkbook = new XSSFWorkbook(inputFileStream);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = inputWorkbook.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet("Validation Results");

            Row inputHeaderRow = inputSheet.getRow(0);
            Row outputHeaderRow = outputSheet.createRow(0);

            // Copy header row to the output file and add "Validation Result" column
            for (int i = 0; i < inputHeaderRow.getLastCellNum(); i++) {
                Cell inputCell = inputHeaderRow.getCell(i);
                Cell outputCell = outputHeaderRow.createCell(i);
                if (inputCell != null) {
                    outputCell.setCellValue(inputCell.getStringCellValue());
                }
            }
            outputHeaderRow.createCell(inputHeaderRow.getLastCellNum()).setCellValue("Validation Result");

            int outputRowIndex = 1;

            // Process each row
            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row inputRow = inputSheet.getRow(i);
                if (inputRow == null) continue;

                String billingCode = getCellValue(inputRow.getCell(findColumnIndex(inputHeaderRow, "BIILING_CODE")));
                List<Map<String, String>> ruleSets = rules.getOrDefault(billingCode, new ArrayList<>());

                ValidationResult result = validateRowAgainstRules(inputRow, inputHeaderRow, ruleSets);

                Row outputRow = outputSheet.createRow(outputRowIndex++);
                copyRow(inputRow, outputRow);
                outputRow.createCell(inputHeaderRow.getLastCellNum()).setCellValue(result.isValid ? "Correct" : "Wrong");

                if (result.matchedRuleRow != null) {
                    Row ruleRow = outputSheet.createRow(outputRowIndex++);
                    populateRuleRow(ruleRow, inputHeaderRow, result.matchedRuleRow);
                    ruleRow.createCell(inputHeaderRow.getLastCellNum()).setCellValue(result.isValid ? "Correct" : "Wrong");
                }
            }

            outputWorkbook.write(outputStream);
        }
        return outputStream;
    }

    private static ValidationResult validateRowAgainstRules(Row row, Row headerRow, List<Map<String, String>> ruleSets) {
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

            if (isMatch) {
                return new ValidationResult(true, ruleSet);
            }
        }

        if (!ruleSets.isEmpty()) {
            return new ValidationResult(false, ruleSets.get(ruleSets.size() - 1));
        }

        return new ValidationResult(false, null);
    }

    private static void copyRow(Row inputRow, Row outputRow) {
        for (int i = 0; i < inputRow.getLastCellNum(); i++) {
            Cell inputCell = inputRow.getCell(i);
            Cell outputCell = outputRow.createCell(i);
            if (inputCell != null) {
                switch (inputCell.getCellType()) {
                    case STRING:
                        outputCell.setCellValue(inputCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        outputCell.setCellValue(inputCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        outputCell.setCellValue(inputCell.getBooleanCellValue());
                        break;
                    default:
                        outputCell.setCellValue("");
                }
            }
        }
    }

    private static void populateRuleRow(Row ruleRow, Row headerRow, Map<String, String> ruleSet) {
        for (Map.Entry<String, String> rule : ruleSet.entrySet()) {
            int colIndex = findColumnIndex(headerRow, rule.getKey());
            ruleRow.createCell(colIndex).setCellValue(rule.getValue());
        }
    }

    private static boolean validateCellValue(String expectedValue, String actualValue) {
        if (expectedValue.equalsIgnoreCase("Not Used")) return true;
        if (expectedValue.startsWith("<>")) {
            String[] excludedValues = expectedValue.substring(3, expectedValue.length() - 1).split(",");
            return Arrays.stream(excludedValues).noneMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        if (expectedValue.contains(",")) {
            String[] allowedValues = expectedValue.split(",");
            return Arrays.stream(allowedValues).anyMatch(val -> val.trim().equalsIgnoreCase(actualValue));
        }
        return expectedValue.equalsIgnoreCase(actualValue);
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

    private static class ValidationResult {
        boolean isValid;
        Map<String, String> matchedRuleRow;

        ValidationResult(boolean isValid, Map<String, String> matchedRuleRow) {
            this.isValid = isValid;
            this.matchedRuleRow = matchedRuleRow;
        }
    }
}
