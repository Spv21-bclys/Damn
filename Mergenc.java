import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelMerger {

    public static ByteArrayOutputStream mergeExcelFiles(
            ByteArrayOutputStream file1Stream,
            ByteArrayOutputStream file2Stream) throws IOException {

        // Configuration: Define common column and column mappings
        String commonColumn = "ID";
        Map<String, String> file1Columns = Map.of(
                "Name", "First Name",
                "Age", "Person Age"
        );
        Map<String, String> file2Columns = Map.of(
                "Salary", "Annual Salary"
        );

        // Read data from the input streams
        List<Map<String, String>> file1Data = readExcel(new ByteArrayInputStream(file1Stream.toByteArray()));
        List<Map<String, String>> file2Data = readExcel(new ByteArrayInputStream(file2Stream.toByteArray()));

        // Merge the data
        List<Map<String, String>> mergedData = mergeData(file1Data, file2Data, commonColumn, file1Columns, file2Columns);

        // Filter data (Optional: adjust as needed)
        List<Map<String, String>> filteredData = filterData(mergedData, "Annual Salary", "Transaction Missing");

        // Write the merged data to a ByteArrayOutputStream
        return writeExcel(filteredData, commonColumn, file1Columns, file2Columns);
    }

    private static List<Map<String, String>> readExcel(InputStream inputStream) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
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

    private static List<Map<String, String>> mergeData(
            List<Map<String, String>> file1Data,
            List<Map<String, String>> file2Data,
            String commonColumn,
            Map<String, String> file1Columns,
            Map<String, String> file2Columns) {

        Map<String, Map<String, String>> file2DataMap = new HashMap<>();
        for (Map<String, String> row : file2Data) {
            file2DataMap.put(row.get(commonColumn), row);
        }

        List<Map<String, String>> mergedData = new ArrayList<>();
        for (Map<String, String> row : file1Data) {
            String key = row.get(commonColumn);
            if (file2DataMap.containsKey(key)) {
                Map<String, String> mergedRow = new LinkedHashMap<>();
                mergedRow.put(commonColumn, key);

                for (String col : file1Columns.keySet()) {
                    mergedRow.put(file1Columns.get(col), row.get(col));
                }

                Map<String, String> file2Row = file2DataMap.get(key);
                for (String col : file2Columns.keySet()) {
                    mergedRow.put(file2Columns.get(col), file2Row.get(col));
                }

                mergedData.add(mergedRow);
            }
        }
        return mergedData;
    }

    private static List<Map<String, String>> filterData(List<Map<String, String>> data, String column, String value) {
        List<Map<String, String>> filteredData = new ArrayList<>();
        for (Map<String, String> row : data) {
            if (value.equals(row.get(column))) {
                filteredData.add(row);
            }
        }
        return filteredData;
    }

    private static ByteArrayOutputStream writeExcel(
            List<Map<String, String>> filteredData,
            String commonColumn,
            Map<String, String> file1Columns,
            Map<String, String> file2Columns) throws IOException {

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Filtered Data");

            Row headerRow = sheet.createRow(0);
            int colIndex = 0;
            headerRow.createCell(colIndex++).setCellValue(commonColumn);
            for (String colName : file1Columns.values()) {
                headerRow.createCell(colIndex++).setCellValue(colName);
            }
            for (String colName : file2Columns.values()) {
                headerRow.createCell(colIndex++).setCellValue(colName);
            }

            int rowIndex = 1;
            for (Map<String, String> row : filteredData) {
                Row excelRow = sheet.createRow(rowIndex++);
                colIndex = 0;
                for (String value : row.values()) {
                    excelRow.createCell(colIndex++).setCellValue(value);
                }
            }

            workbook.write(outputStream);
        }
        return outputStream;
    }
}
