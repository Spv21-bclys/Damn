import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class MergeExcelByColumnName {

    public static ByteArrayOutputStream addPriorityOrderToExcel(ByteArrayOutputStream firstExcelStream, String secondExcelFilePath) {
        ByteArrayOutputStream updatedFirstExcelStream = new ByteArrayOutputStream();
        
        try {
            // Load the first Excel file from ByteArrayOutputStream
            Workbook workbook1 = new XSSFWorkbook(new ByteArrayInputStream(firstExcelStream.toByteArray()));
            Sheet sheet1 = workbook1.getSheetAt(0);

            // Load the second Excel file from the file path
            FileInputStream file2 = new FileInputStream(secondExcelFilePath);
            Workbook workbook2 = new XSSFWorkbook(file2);
            Sheet sheet2 = workbook2.getSheetAt(0);

            // Get column indices based on column names
            int billingCodeIndexFile1 = getColumnIndexByName(sheet1, "BillingCode");
            int billingCodeIndexFile2 = getColumnIndexByName(sheet2, "BillingCode");
            int priorityOrderIndex = getColumnIndexByName(sheet2, "PriorityOrder");

            if (billingCodeIndexFile1 == -1 || billingCodeIndexFile2 == -1 || priorityOrderIndex == -1) {
                System.out.println("Column names not found in one of the sheets.");
                return null;
            }

            // Create a map for BillingCode -> PriorityOrder from the second Excel file
            Map<String, String> billingCodeToPriorityOrder = new HashMap<>();
            for (Row row : sheet2) {
                Cell billingCodeCell = row.getCell(billingCodeIndexFile2);
                Cell priorityOrderCell = row.getCell(priorityOrderIndex);

                if (billingCodeCell != null && priorityOrderCell != null) {
                    billingCodeToPriorityOrder.put(billingCodeCell.toString(), priorityOrderCell.toString());
                }
            }

            // Shift columns to the right to make space for PriorityOrder beside BillingCode
            for (Row row : sheet1) {
                if (row.getLastCellNum() > billingCodeIndexFile1) {
                    for (int col = row.getLastCellNum(); col > billingCodeIndexFile1; col--) {
                        Cell oldCell = row.getCell(col - 1);
                        Cell newCell = row.createCell(col);

                        if (oldCell != null) {
                            newCell.setCellValue(oldCell.toString());
                        } else {
                            row.removeCell(newCell);
                        }
                    }
                }
            }

            // Add PriorityOrder header beside BillingCode
            Row headerRow = sheet1.getRow(0);
            if (headerRow != null) {
                headerRow.createCell(billingCodeIndexFile1 + 1).setCellValue("PriorityOrder");
            }

            // Add PriorityOrder values to the rows
            for (Row row : sheet1) {
                Cell billingCodeCell = row.getCell(billingCodeIndexFile1);

                if (billingCodeCell != null) {
                    String billingCodeValue = billingCodeCell.toString();
                    String priorityOrderValue = billingCodeToPriorityOrder.get(billingCodeValue);

                    if (priorityOrderValue != null) {
                        Cell priorityOrderCell = row.createCell(billingCodeIndexFile1 + 1);
                        priorityOrderCell.setCellValue(priorityOrderValue);
                    }
                }
            }

            // Write the updated workbook back to ByteArrayOutputStream
            workbook1.write(updatedFirstExcelStream);

            // Close workbooks
            workbook1.close();
            workbook2.close();

            System.out.println("PriorityOrder column added beside BillingCode in the first Excel file successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        return updatedFirstExcelStream;
    }

    // Helper method to get the column index by name
    private static int getColumnIndexByName(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0); // Assuming the first row is the header
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Return -1 if the column name is not found
    }
}
