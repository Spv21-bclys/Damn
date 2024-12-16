import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelValidatorWithMultipleValues {

    public static void main(String[] args) throws IOException {
        String rulesFilePath = "ExcelA.xlsx"; // Replace with your rules file
        String dataFilePath = "ExcelB.xlsx";  // Replace with your data file

        // Read rules from Excel A
        Map<String, Map<String, List<String>>> rulesMap = readRules(rulesFilePath);

        // Validate Excel B against the rules and update the same file
        validateAndWriteResults(dataFilePath, rulesMap);

        System.out.println("Validation complete. Results added to ExcelB.xlsx");
    }

    private static Map<String, Map<String, List<String>>> readRules(String filePath) throws IOException {
        Map<String, Map<String, List<String>>> rules = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            // List of column names from the rule book
            List<String> columnNames = Arrays.asList(
                    "BIILING_CODE", "CHARGING_INIDICATOR", "CURRENCY", "FINAL MOP", "RECEIVER_BIC",
                    "PSD_INDICATOR", "PYMT_DEST_CTRY", "SWIFT_MSG_TYP", "DR_TRN_CODES", "CR_TRN_CODES", "FI_CHARGING_INDICATOR"
            );

            // Initialize rules for each column
            for (String columnName : columnNames) {
                int colIndex = findColumnIndex(headerRow, columnName);
                rules.putIfAbsent(columnName, new HashMap<>());

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String checkValue = getCellValue(row.getCell(colIndex));  // Get check value safely
                    String columnValue = getCellValue(row.getCell(colIndex));  // Get column value safely

                    // Handle case where value is not used
                    if ("Not Used".equalsIgnoreCase(columnValue)) {
                        continue;
                    }
                    // Handle case where value should be excluded (e.g., <> (v,b))
                    if (columnValue.startsWith("<>")) {
                        List<String> excludedValues = parseExcludedValues(columnValue);
                        rules.get(columnName).put(checkValue, excludedValues);
                    } else {
                        rules.get(columnName).putIfAbsent(columnValue, new ArrayList<>());
                    }
                }
            }
        }
        return rules;
    }

    private static List<String> parseExcludedValues(String value) {
        // Example: <> (af,sd) -> return [af, sd]
        String excludedValuesString = value.substring(3, value.length() - 1); // Remove <> and parentheses
        return Arrays.asList(excludedValuesString.split(","));
    }

    private static void validateAndWriteResults(String filePath, Map<String, Map<String, List<String>>> rules) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(filePath)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            // Add a new column for the validation result
            int resultColIndex = headerRow.getPhysicalNumberOfCells();
            headerRow.createCell(resultColIndex).setCellValue("Result");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                boolean isValid = true;  // Initially assume the row is valid
                StringBuilder validationResult = new StringBuilder();

                // Validate each column using the rules from Excel A
                for (String columnName : rules.keySet()) {
                    String columnValue = getCellValue(row.getCell(findColumnIndex(headerRow, columnName)));  // Get column value safely
                    Map<String, List<String>> validValues = rules.get(columnName);

                    // Handle multiple values in a cell
                    List<String> valuesToValidate = Arrays.asList(columnValue.split(","));
                    boolean columnValid = true;

                    // Check if at least one value is incorrect
                    for (String value : valuesToValidate) {
                        if (!validateColumn(value.trim(), validValues, columnName)) {
                            columnValid = false;
                            break;  // If one value is invalid, we can stop checking
                        }
                    }

                    if (!columnValid) {
                        isValid = false;  // If any column is invalid, the whole row is invalid
                        validationResult.append("Wrong;");
                    } else {
                        validationResult.append("Correct;");
                    }
                }

                // Set the validation result in the new column
                row.createCell(resultColIndex).setCellValue(isValid ? "Correct" : "Wrong");
            }

            // Write the updated workbook to the same file
            workbook.write(fos);
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private static boolean validateColumn(String columnValue, Map<String, List<String>> validValues, String columnName) {
        // Handle "Not Used" Case: Any value is valid
        if ("Not Used".equalsIgnoreCase(columnValue)) {
            return true;
        }

        // Handle "<>" Exclusion Rule
        List<String> excludedValues = validValues.get(columnName);
        if (excludedValues != null && !excludedValues.isEmpty()) {
            return !excludedValues.contains(columnValue);  // Valid if the value is not in the exclusion list
        }

        // If it's a valid value in the rule book
        return validValues.containsKey(columnValue);
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
