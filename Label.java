import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelValidator {

    public static void main(String[] args) {
        try {
            // File paths
            String ruleBookPath = "ruleBook.xlsx";
            String targetFilePath = "target.xlsx";
            String outputPath = "validatedOutput.xlsx";

            // Read Excel files
            List<Map<String, String>> ruleBookData = readExcel(ruleBookPath);
            List<Map<String, String>> targetData = readExcel(targetFilePath);

            // Validate data
            List<Map<String, String>> validatedData = validateData(ruleBookData, targetData);

            // Write output to a new Excel file
            writeExcel(validatedData, outputPath);

            System.out.println("Validation completed successfully. Output saved to: " + outputPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Reads an Excel file and returns data as a list of maps (each map represents a row)
    private static List<Map<String, String>> readExcel(String filePath) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Read the header row
            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            // Read the remaining rows
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Map<String, String> rowData = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData.put(headers.get(i), getCellValueAsString(cell));
                }
                data.add(rowData);
            }
        }
        return data;
    }

    // Converts a cell to a string value
    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
            default:
                return "";
        }
    }

    // Validates target data against the rule book
    private static List<Map<String, String>> validateData(List<Map<String, String>> ruleBookData, List<Map<String, String>> targetData) {
        Map<String, Map<String, String>> ruleMap = new HashMap<>();

        // Build a map for quick access to rules by BIILING_CODE
        for (Map<String, String> ruleRow : ruleBookData) {
            String billingCode = ruleRow.get("BIILING_CODE");
            ruleMap.put(billingCode, ruleRow);
        }

        // Validate target data
        for (Map<String, String> targetRow : targetData) {
            String billingCode = targetRow.get("BIILING_CODE");
            Map<String, String> rules = ruleMap.get(billingCode);

            if (rules != null) {
                boolean isValid = true;
                for (String column : rules.keySet()) {
                    if (!column.equals("BIILING_CODE")) {
                        String ruleValue = rules.get(column);
                        String targetValue = targetRow.get(column);

                        if (!matchesRule(ruleValue, targetValue)) {
                            isValid = false;
                            break;
                        }
                    }
                }
                targetRow.put("Label", isValid ? "Correct" : "Wrong");
            } else {
                targetRow.put("Label", "Wrong");
            }
        }

        return targetData;
    }

    // Checks if a target value matches a rule
    private static boolean matchesRule(String rule, String value) {
        if (rule.equals("Not Used")) {
            return true; // Any value is valid
        } else if (rule.startsWith("<>") && rule.contains(",")) {
            // Handle rules like <>(v,n) or <>(b,c)
            String[] invalidValues = rule.substring(2).split(",");
            for (String invalid : invalidValues) {
                if (value.equals(invalid.trim())) {
                    return false;
                }
            }
            return true;
        } else if (rule.contains(",")) {
            // Handle multiple valid values like G,H
            String[] validValues = rule.split(",");
            for (String valid : validValues) {
                if (value.equals(valid.trim())) {
                    return true;
                }
            }
            return false;
        } else {
            // Handle single value
            return rule.equals(value);
        }
    }

    // Writes validated data to an Excel file
    private static void writeExcel(List<Map<String, String>> data, String outputPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(new File(outputPath))) {

            Sheet sheet = workbook.createSheet("Validated Data");

            // Write header row
            Row headerRow = sheet.createRow(0);
            Set<String> headers = data.get(0).keySet();
            int colIndex = 0;
            for (String header : headers) {
                headerRow.createCell(colIndex++).setCellValue(header);
            }

            // Write data rows
            int rowIndex = 1;
            for (Map<String, String> row : data) {
                Row excelRow = sheet.createRow(rowIndex++);
                colIndex = 0;
                for (String value : row.values()) {
                    excelRow.createCell(colIndex++).setCellValue(value);
                }
            }

            workbook.write(fos);
        }
    }
}
