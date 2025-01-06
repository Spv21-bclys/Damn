import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelMerger {
    public static void main(String[] args) {
        String file1Path = "excel1.xlsx"; // Replace with your file path
        String file2Path = "excel2.xlsx"; // Replace with your file path
        String outputFilePath = "merged_output.xlsx";

        try {
            // Read both Excel files
            Workbook workbook1 = new XSSFWorkbook(new FileInputStream(file1Path));
            Workbook workbook2 = new XSSFWorkbook(new FileInputStream(file2Path));

            // Combine the sheets into a single workbook
            Workbook mergedWorkbook = mergeExcelFiles(workbook1, workbook2);

            // Write the merged workbook to a new file
            FileOutputStream outputStream = new FileOutputStream(outputFilePath);
            mergedWorkbook.write(outputStream);

            outputStream.close();
            System.out.println("Merged file created: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Workbook mergeExcelFiles(Workbook workbook1, Workbook workbook2) {
        Workbook mergedWorkbook = new XSSFWorkbook();
        Sheet sheet1 = workbook1.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(0);

        // Create a new sheet in the merged workbook
        Sheet mergedSheet = mergedWorkbook.createSheet("MergedSheet");

        // Map to store unique headers
        LinkedHashMap<String, Integer> headers = new LinkedHashMap<>();

        // Helper to track the current row
        int currentRow = 0;

        // Add headers and data from the first sheet
        currentRow = copyHeadersAndData(sheet1, mergedSheet, headers, currentRow);

        // Add headers and data from the second sheet (excluding a specific column)
        currentRow = copyHeadersAndDataExcludingColumn(sheet2, mergedSheet, headers, currentRow, "ColumnToExclude");

        return mergedWorkbook;
    }

    public static int copyHeadersAndData(Sheet sourceSheet, Sheet targetSheet,
                                         LinkedHashMap<String, Integer> headers, int currentRow) {
        int lastColumn = sourceSheet.getRow(0).getLastCellNum();

        // Add headers if they are new
        Row headerRow = targetSheet.getRow(0) == null ? targetSheet.createRow(0) : targetSheet.getRow(0);
        for (int i = 0; i < lastColumn; i++) {
            Cell cell = sourceSheet.getRow(0).getCell(i);
            String header = cell.getStringCellValue();

            if (!headers.containsKey(header)) {
                headers.put(header, headers.size());
                headerRow.createCell(headers.get(header)).setCellValue(header);
            }
        }

        // Add data
        for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row targetRow = targetSheet.getRow(currentRow) == null ? targetSheet.createRow(currentRow) : targetSheet.getRow(currentRow);

            for (int j = 0; j < lastColumn; j++) {
                Cell sourceCell = sourceRow.getCell(j);
                if (sourceCell != null) {
                    String header = sourceSheet.getRow(0).getCell(j).getStringCellValue();
                    int columnIndex = headers.get(header);

                    Cell targetCell = targetRow.createCell(columnIndex);
                    copyCellValue(sourceCell, targetCell);
                }
            }
            currentRow++;
        }
        return currentRow;
    }

    public static int copyHeadersAndDataExcludingColumn(Sheet sourceSheet, Sheet targetSheet,
                                                        LinkedHashMap<String, Integer> headers, int currentRow, String columnToExclude) {
        int lastColumn = sourceSheet.getRow(0).getLastCellNum();

        // Add headers and skip the excluded column
        Row headerRow = targetSheet.getRow(0) == null ? targetSheet.createRow(0) : targetSheet.getRow(0);
        for (int i = 0; i < lastColumn; i++) {
            Cell cell = sourceSheet.getRow(0).getCell(i);
            String header = cell.getStringCellValue();

            // Skip the excluded column
            if (!header.equals(columnToExclude) && !headers.containsKey(header)) {
                headers.put(header, headers.size());
                headerRow.createCell(headers.get(header)).setCellValue(header);
            }
        }

        // Add data, skipping the excluded column
        for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row targetRow = targetSheet.getRow(currentRow) == null ? targetSheet.createRow(currentRow) : targetSheet.getRow(currentRow);

            for (int j = 0; j < lastColumn; j++) {
                Cell sourceCell = sourceRow.getCell(j);
                String header = sourceSheet.getRow(0).getCell(j).getStringCellValue();

                // Skip the excluded column
                if (sourceCell != null && !header.equals(columnToExclude)) {
                    int columnIndex = headers.get(header);
                    Cell targetCell = targetRow.createCell(columnIndex);
                    copyCellValue(sourceCell, targetCell);
                }
            }
            currentRow++;
        }
        return currentRow;
    }

    public static void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                targetCell.setBlank();
        }
    }
}
