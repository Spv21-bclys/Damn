import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class ExcelMapper {
    
    private static Map<String, Integer> getColumnIndexMap(Sheet sheet) {
        Map<String, Integer> columnIndexMap = new HashMap<>();
        Row headerRow = sheet.getRow(0); // Assuming first row is the header
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                columnIndexMap.put(cell.getStringCellValue(), cell.getColumnIndex());
            }
        }
        return columnIndexMap;
    }

    public static void mapAndSaveExcel(String excel1Path, String excel2Path, String outputPath) throws IOException {
        FileInputStream file1 = new FileInputStream(new File(excel1Path));
        FileInputStream file2 = new FileInputStream(new File(excel2Path));
        Workbook workbook1 = new XSSFWorkbook(file1);
        Workbook workbook2 = new XSSFWorkbook(file2);

        Sheet sheet1 = workbook1.getSheetAt(0); // Excel 1
        Sheet sheet2 = workbook2.getSheetAt(0); // Excel 2

        // Get column indices dynamically
        Map<String, Integer> indexMap1 = getColumnIndexMap(sheet1);
        Map<String, Integer> indexMap2 = getColumnIndexMap(sheet2);

        // Validate required columns
        if (!indexMap1.containsKey("PRICE_ITEM") || !indexMap2.containsKey("PRICE_ITEM") || !indexMap2.containsKey("Result Set CSV")) {
            System.out.println("Missing required columns in one of the files.");
            return;
        }

        int colC1 = indexMap1.get("PRICE_ITEM");
        int colC2 = indexMap2.get("PRICE_ITEM");
        int colD2 = indexMap2.get("Result Set CSV");

        // Create a mapping of PRICE_ITEM -> Result Set CSV from Excel 2
        Map<String, String> mapping = new HashMap<>();
        for (Row row : sheet2) {
            Cell cellC = row.getCell(colC2);
            Cell cellD = row.getCell(colD2);
            if (cellC != null && cellD != null) {
                mapping.put(cellC.toString(), cellD.toString());
            }
        }

        // Create output workbook
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Mapped Data");

        // Write headers
        Row header = outputSheet.createRow(0);
        header.createCell(0).setCellValue("PNCSREFNUM");
        header.createCell(1).setCellValue("TXNACCNUM");
        header.createCell(2).setCellValue("PRICE_ITEM");
        header.createCell(3).setCellValue("Result Set CSV");

        // Process Excel 1 and map values
        int rowNum = 1;
        for (Row row : sheet1) {
            Cell cellC = row.getCell(colC1);
            if (cellC != null && mapping.containsKey(cellC.toString())) {
                Row newRow = outputSheet.createRow(rowNum++);
                newRow.createCell(0).setCellValue(row.getCell(indexMap1.get("PNCSREFNUM")).toString()); // Column A
                newRow.createCell(1).setCellValue(row.getCell(indexMap1.get("TXNACCNUM")).toString()); // Column B
                newRow.createCell(2).setCellValue(cellC.toString());                           // Column C
                newRow.createCell(3).setCellValue(mapping.get(cellC.toString()));             // Column D
            }
        }

        // Save output file
        FileOutputStream fileOut = new FileOutputStream(outputPath);
        outputWorkbook.write(fileOut);
        fileOut.close();
        workbook1.close();
        workbook2.close();
        outputWorkbook.close();

        System.out.println("Merged file saved as " + outputPath);
    }

    public static void main(String[] args) throws IOException {
        mapAndSaveExcel("excel1.xlsx", "excel2.xlsx", "output.xlsx");
    }
}
