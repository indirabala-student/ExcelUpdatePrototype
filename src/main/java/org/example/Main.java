package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        System.out.println("Starting Excel Processing...");

        String staticFilePath= "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\RCPLNP01.xlsx";
        String dynamicFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\example.xlsx";
        String outputFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\out_data.xlsx";
        String flagColumn = "Flag"; // The column that determines if updates are needed

        try (FileInputStream fisStatic = new FileInputStream(staticFilePath);
             FileInputStream fisDynamic = new FileInputStream(dynamicFilePath);
             Workbook workbookStatic = new XSSFWorkbook(fisStatic);
             Workbook workbookDynamic= new XSSFWorkbook(fisDynamic)) {

            Sheet staticSheet = workbookStatic.getSheetAt(0); // Sheet 1 (Reference)
            Sheet dynamicSheet = workbookDynamic.getSheet("000A"); // Sheet 2 (To be updated)

            // Get column headers for both sheets
            List<String> headerStaticSheet = getColumnHeaders(staticSheet);
            List<String> headerDynamicSheet = getColumnHeaders(dynamicSheet);

            // Find common columns between both sheets
            Set<String> commonColumns = new HashSet<>(headerStaticSheet);
            commonColumns.retainAll(headerDynamicSheet);

            // Get column index mappings for both sheets
            Map<String, Integer> indexMapStatic = getColumnIndexMap(staticSheet);
            Map<String, Integer> indexMapDynamic = getColumnIndexMap(dynamicSheet);

            // Validate if flag column exists in Sheet 2
            if (!indexMapDynamic.containsKey(flagColumn)) {
                System.err.println("Error: Flag column '" + flagColumn + "' not found in Sheet 2.");
                return;
            }
            int flagColumnIndex = indexMapDynamic.get(flagColumn);

            // Process Sheet 2 rows for updating common columns
            for (int i = 1; i < dynamicSheet.getPhysicalNumberOfRows(); i++) {
                Row rowDynamic = dynamicSheet.getRow(i);
                Row rowStatic = staticSheet.getRow(i); // Match row with the same index

                if (rowDynamic == null || rowStatic == null) {
                    continue; // Skip empty rows
                }

                // Get flag column value in Sheet 2
                Cell flagCell = rowDynamic.getCell(flagColumnIndex);
                if (flagCell != null && "Y".equalsIgnoreCase(flagCell.getStringCellValue().trim())) {
                    // Update all common columns from Sheet 1 to Sheet 2
                    for (String column : commonColumns) {
                        int indexStatic = indexMapStatic.get(column);
                        int indexDynamic = indexMapDynamic.get(column);

                        Cell cellStatic = rowStatic.getCell(indexStatic);
                        Cell cellDynamic = rowDynamic.getCell(indexDynamic);

                        if (cellStatic != null) {
                            if (cellDynamic == null) {
                                cellDynamic = rowDynamic.createCell(indexDynamic); // Create cell if missing
                            }
                            cellDynamic.setCellValue(cellStatic.toString()); // Copy value from Sheet 1
                        }
                    }
                }
            }

            // Save the updated Sheet 2 back to an Excel file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbookDynamic.write(fos);
            }

            System.out.println("Dynamic Sheet updated successfully!");

        } catch (IOException e) {
            System.err.println("Error processing the Excel file: " + e.getMessage());
        }
    }

    /**
     * Retrieves a map of column names to their index positions.
     */
    private static Map<String, Integer> getColumnIndexMap(Sheet sheet) {
        Map<String, Integer> indexMap = new HashMap<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
                indexMap.put(headerRow.getCell(i).getStringCellValue().trim(), i);
            }
        }
        return indexMap;
    }

    /**
     * Retrieves a list of column headers from the sheet.
     */
    public static List<String> getColumnHeaders(Sheet sheet) {
        List<String> headers = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue().trim());
            }
        }
        return headers;
    }

    /**
     * Reads and prints the data from a sheet (for debugging).
     */
    public static void readSheetData(Sheet sheet) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                System.out.print(cell.toString() + "\t");
            }
            System.out.println();
        }
    }
}
