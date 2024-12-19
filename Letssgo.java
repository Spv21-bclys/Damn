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
            validateAndUpdateExcel(inputFilePath, outputFilePath, rules);
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

    private static void validateAndUpdateExcel(String inputFilePath, String outputFilePath,
                                               Map<String, List<Map<String, String>>> rules) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            CellStyle redStyle = createRedCellStyle(workbook); // Reuse this style

            // Process each row
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String billingCode = getCellValue(row.getCell(findColumnIndex(headerRow, "BIILING_CODE")));
                List<Map<String, String>> ruleSets = rules.getOrDefault(billingCode, new ArrayList<>());

                boolean isMatched = false;
                Map<String, String> matchedRule = null;

                for (Map<String, String> ruleSet : ruleSets) {
                    boolean isValid = validateRowAgainstRule(row, headerRow, ruleSet, redStyle);
                    if (isValid) {
                        matchedRule = ruleSet;
                        isMatched = true;
                        break;
                    }
                }

                if (!isMatched && !ruleSets.isEmpty()) {
                    matchedRule = ruleSets.get(0); // Use the first rule as fallback
                }

                // Append only if there is a valid rule to add
                if (matchedRule != null) {
                    appendRuleToSheet(sheet, matchedRule);
                }
            }

            // Save the updated Excel file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        }
    }

    private static boolean validateRowAgainstRule(Row row, Row headerRow, Map<String, String> ruleSet, CellStyle redStyle) {
        boolean isMatch = true;

        for (Map.Entry<String, String> rule : ruleSet.entrySet()) {
            String columnName = rule.getKey();
            String expectedValue = rule.getValue();
            int colIndex = findColumnIndex(headerRow, columnName);
            Cell cell = row.getCell(colIndex);

            String actualValue = getCellValue(cell);
            if (!validateCellValue(expectedValue, actualValue)) {
                isMatch = false;
                if (cell != null) {
                    cell.setCellStyle(redStyle); // Highlight mismatched cells
                }
            }
        }
        return isMatch;
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

    private static void appendRuleToSheet(Sheet sheet, Map<String, String> ruleSet) {
        Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
        int colIndex = 0;
        for (String value : ruleSet.values()) {
            Cell cell = newRow.createCell(colIndex++);
            cell.setCellValue(value);
        }
    }

    private static CellStyle createRedCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        return style;
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
