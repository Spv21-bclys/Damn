import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelValidatorWithAdditionalRules {

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

            int checkCol = findColumnIndex(headerRow, "Check");
            int col1Index = findColumnIndex(headerRow, "Column1");
            int col2Index = findColumnIndex(headerRow, "Column2");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String checkValue = row.getCell(checkCol).getStringCellValue();
                String column1Value = row.getCell(col1Index).getStringCellValue();
                String column2Value = row.getCell(col2Index).getStringCellValue();

                // Initialize map for each "Check" value
                rules.putIfAbsent(checkValue, new HashMap<>());

                // Add values for Column1
                rules.get(checkValue).putIfAbsent("Column1", new ArrayList<>());
                rules.get(checkValue).get("Column1").add(column1Value);

                // Add values for Column2 (split by comma for multiple values)
                rules.get(checkValue).putIfAbsent("Column2", new ArrayList<>());
                for (String value : column2Value.split(",")) {
                    rules.get(checkValue).get("Column2").add(value.trim());
                }
            }
        }
        return rules;
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

            int checkCol = findColumnIndex(headerRow, "Check");
            int col1Index = findColumnIndex(headerRow, "Column1");
            int col2Index = findColumnIndex(headerRow, "Column2");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String checkValue = row.getCell(checkCol).getStringCellValue();
                String column1Value = row.getCell(col1Index).getStringCellValue();
                String column2Value = row.getCell(col2Index).getStringCellValue();

                // Validate against rules
                String result;
                if (rules.containsKey(checkValue)) {
                    Map<String, List<String>> validValues = rules.get(checkValue);
                    boolean isValid = validateColumn1(column1Value, validValues.get("Column1")) &&
                                      validValues.get("Column2").contains(column2Value);
                    result = isValid ? "Correct" : "Wrong";
                } else {
                    result = "Wrong";
                }

                // Set the validation result in the new column
                row.createCell(resultColIndex).setCellValue(result);
            }

            // Write the updated workbook to the same file
            workbook.write(fos);
        }
    }

    private static boolean validateColumn1(String column1Value, List<String> validValues) {
        // Check for "Not Used" case
        if (validValues.contains("Not Used")) {
            return true;
        }

        // Check for "<>" case (values to exclude)
        for (String valid : validValues) {
            if (valid.startsWith("<>")) {
                List<String> excludedValues = parseExcludedValues(valid);
                if (excludedValues.contains(column1Value)) {
                    return false; // Column1 value is in excluded list, hence invalid
                }
            } else {
                if (valid.equals(column1Value)) {
                    return true;
                }
            }
        }

        return false; // If no match found, it's invalid
    }

    private static List<String> parseExcludedValues(String rule) {
        // Extract values inside <> and split by comma
        String values = rule.substring(2, rule.length() - 1);  // Remove "<>" part
        return Arrays.asList(values.split(","));
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
