import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelReader {

    public static void main(String[] args) {
        try {
            String filePath = "example.xlsx";  // Path to the Excel file
            List<Map<String, String>> excelData = readExcel(filePath);
            // Print the read data
            for (Map<String, String> row : excelData) {
                System.out.println(row);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Function to read Excel file and return data as a list of maps (each map is a row)
    public static List<Map<String, String>> readExcel(String filePath) throws IOException, InvalidFormatException {
        List<Map<String, String>> data = new ArrayList<>();
        
        // Open the Excel file
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Ensure the sheet is not empty
            if (!rowIterator.hasNext()) {
                throw new IllegalArgumentException("The sheet is empty.");
            }

            // Read the header row (first row) and store column names
            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue()); // Get column names from the header row
            }

            // Read the data rows
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Map<String, String> rowData = new HashMap<>();

                // Iterate through each cell in the row and map it to the corresponding column name
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData.put(headers.get(i), getCellValueAsString(cell)); // Add data to map with header as key
                }

                // Add the row map to the data list
                data.add(rowData);
            }
        }

        return data;  // Return the list of maps containing the rows of the Excel sheet
    }

    // Helper function to get the value of a cell as a string
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
            default:
                return "";
        }
    }
}
