import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelMerger {

    public static void main(String[] args) {
        try {
            // Define input files and output file paths
            String file1Path = "file1.xlsx";
            String file2Path = "file2.xlsx";
            String outputPath = "output.xlsx";

            // Define the common column to merge on
            String commonColumn = "ID";

            // Define the required columns and their new names
            Map<String, String> file1Columns = Map.of(
                "Name", "First Name",
                "Age", "Person Age"
            );

            Map<String, String> file2Columns = Map.of(
                "Salary", "Annual Salary"
            );

            // Read Excel files
            List<Map<String, String>> file1Data = readExcel(file1Path);
            List<Map<String, String>> file2Data = readExcel(file2Path);

            // Merge data
            List<Map<String, String>> mergedData = mergeData(file1Data, file2Data, commonColumn, file1Columns, file2Columns);

            // Write output to a new Excel file
            writeExcel(mergedData, commonColumn, file1Columns, file2Columns, outputPath);

            System.out.println("Data merged successfully. Output saved to: " + outputPath);

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

    // Merges data from two files based on a common column and required columns
    private static List<Map<String, String>> mergeData(
            List<Map<String, String>> file1Data,
            List<Map<String, String>> file2Data,
            String commonColumn,
            Map<String, String> file1Columns,
            Map<String, String> file2Columns) {

        // Create a map for quick lookup of file2 rows by the common column
        Map<String, Map<String, String>> file2DataMap = new HashMap<>();
        for (Map<String, String> row : file2Data) {
            file2DataMap.put(row.get(commonColumn), row);
        }

        List<Map<String, String>> mergedData = new ArrayList<>();

        // Merge rows
        for (Map<String, String> row : file1Data) {
            String key = row.get(commonColumn);
            if (file2DataMap.containsKey(key)) {
                Map<String, String> mergedRow = new LinkedHashMap<>();
                mergedRow.put(commonColumn, key);

                // Add required columns from file1
                for (String col : file1Columns.keySet()) {
                    mergedRow.put(file1Columns.get(col), row.get(col));
                }

                // Add required columns from file2
                Map<String, String> file2Row = file2DataMap.get(key);
                for (String col : file2Columns.keySet()) {
                    mergedRow.put(file2Columns.get(col), file2Row.get(col));
                }

                mergedData.add(mergedRow);
            }
        }
        return mergedData;
    }

    // Writes merged data to an Excel file
    private static void writeExcel(
            List<Map<String, String>> mergedData,
            String commonColumn,
            Map<String, String> file1Columns,
            Map<String, String> file2Columns,
            String outputPath) throws IOException {

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(new File(outputPath))) {

            Sheet sheet = workbook.createSheet("Merged Data");

            // Write header row
            Row headerRow = sheet.createRow(0);
            int colIndex = 0;
            headerRow.createCell(colIndex++).setCellValue(commonColumn);
            for (String colName : file1Columns.values()) {
                headerRow.createCell(colIndex++).setCellValue(colName);
            }
            for (String colName : file2Columns.values()) {
                headerRow.createCell(colIndex++).setCellValue(colName);
            }

            // Write data rows
            int rowIndex = 1;
            for (Map<String, String> row : mergedData) {
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
