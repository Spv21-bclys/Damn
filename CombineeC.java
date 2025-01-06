import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelMerger {

    public static ByteArrayOutputStream mergeExcelFiles(ByteArrayOutputStream baos1, ByteArrayOutputStream baos2, ByteArrayOutputStream baos3) throws IOException {
        // Define column names and configuration
        String commonColumn = "C"; // Shared column
        String columnFromFile1 = "Column1_from_file1"; // Specific column from file1
        List<String> columnsFromFile2 = Arrays.asList("ColumnA", "ColumnB", "Additional1", "Additional2", "Additional3");
        List<String> columnsFromFile3 = Arrays.asList("ColumnA", "ColumnB", "UniqueTo3");

        // Load data from ByteArrayOutputStream
        Map<String, List<String>> data1 = loadExcelData(baos1, commonColumn, Collections.singletonList(columnFromFile1));
        Map<String, List<String>> data2 = loadExcelData(baos2, commonColumn, columnsFromFile2);
        Map<String, List<String>> data3 = loadExcelData(baos3, commonColumn, columnsFromFile3);

        // Combine data
        Map<String, List<String>> combinedData = new LinkedHashMap<>();
        Set<String> allKeys = new HashSet<>();
        allKeys.addAll(data1.keySet());
        allKeys.addAll(data2.keySet());
        allKeys.addAll(data3.keySet());

        for (String key : allKeys) {
            List<String> combinedRow = new ArrayList<>();
            combinedRow.add(key); // Add common column
            combinedRow.addAll(data1.getOrDefault(key, Arrays.asList("", ""))); // Add data from file1
            combinedRow.addAll(data2.getOrDefault(key, new ArrayList<>(Collections.nCopies(columnsFromFile2.size(), "")))); // Add data from file2
            combinedRow.addAll(data3.getOrDefault(key, new ArrayList<>(Collections.nCopies(columnsFromFile3.size(), "")))); // Add data from file3
            combinedData.put(key, combinedRow);
        }

        // Write combined data to ByteArrayOutputStream
        return writeCombinedDataToByteArray(combinedData, columnFromFile1, columnsFromFile2, columnsFromFile3);
    }

    private static Map<String, List<String>> loadExcelData(ByteArrayOutputStream baos, String commonColumn, List<String> specificColumns) throws IOException {
        Map<String, List<String>> data = new LinkedHashMap<>();

        try (ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
             Workbook workbook = new XSSFWorkbook(bais)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            // Map column names to their indexes
            Map<String, Integer> columnIndexes = new HashMap<>();
            for (Cell cell : headerRow) {
                columnIndexes.put(cell.getStringCellValue(), cell.getColumnIndex());
            }

            if (!columnIndexes.containsKey(commonColumn)) {
                throw new IllegalArgumentException("Common column '" + commonColumn + "' not found in input data.");
            }

            List<Integer> selectedIndexes = new ArrayList<>();
            selectedIndexes.add(columnIndexes.get(commonColumn)); // Always include the common column

            for (String column : specificColumns) {
                if (columnIndexes.containsKey(column)) {
                    selectedIndexes.add(columnIndexes.get(column));
                } else {
                    System.out.println("Warning: Column '" + column + "' not found in input data.");
                }
            }

            // Read data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    String key = row.getCell(columnIndexes.get(commonColumn)).toString();
                    List<String> rowData = new ArrayList<>();
                    for (int index : selectedIndexes) {
                        Cell cell = row.getCell(index);
                        rowData.add(cell != null ? cell.toString() : "");
                    }
                    data.put(key, rowData);
                }
            }
        }

        return data;
    }

    private static ByteArrayOutputStream writeCombinedDataToByteArray(
            Map<String, List<String>> data,
            String columnFromFile1,
            List<String> columnsFromFile2,
            List<String> columnsFromFile3) throws IOException {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Combined Data");

            // Write header
            Row headerRow = sheet.createRow(0);
            int colIndex = 0;
            headerRow.createCell(colIndex++).setCellValue("Common_Column");
            headerRow.createCell(colIndex++).setCellValue(columnFromFile1);
            for (String col : columnsFromFile2) {
                headerRow.createCell(colIndex++).setCellValue(col);
            }
            for (String col : columnsFromFile3) {
                headerRow.createCell(colIndex++).setCellValue(col);
            }

            // Write data rows
            int rowIndex = 1;
            for (List<String> rowData : data.values()) {
                Row row = sheet.createRow(rowIndex++);
                for (int cellIndex = 0; cellIndex < rowData.size(); cellIndex++) {
                    row.createCell(cellIndex).setCellValue(rowData.get(cellIndex));
                }
            }

            // Autosize columns
            for (int col = 0; col < headerRow.getLastCellNum(); col++) {
                sheet.autoSizeColumn(col);
            }

            // Write to ByteArrayOutputStream
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos;
        }
    }
}
