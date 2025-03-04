package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        System.out.println("Starting Excel Processing...");

        String staticFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\RCPLNP01.xlsx";
        String dynamicFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\example.xlsx";
        String outputFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\out_data.xlsx";
        String flagColumn = "Flag"; // Column that determines if updates are needed

        try (FileInputStream fisStatic = new FileInputStream(staticFilePath);
             FileInputStream fisDynamic = new FileInputStream(dynamicFilePath);
             Workbook workbookStatic = new XSSFWorkbook(fisStatic);
             Workbook workbookDynamic = new XSSFWorkbook(fisDynamic)) {

            Sheet staticSheet = workbookStatic.getSheetAt(0); // Reference Sheet
            Sheet dynamicSheet = workbookDynamic.getSheet("000A"); // Sheet to update

            // Get column headers from correct rows
            List<String> headerStaticSheet = getColumnHeaders(staticSheet, 0); // 1st row in Excel
            List<String> headerDynamicSheet = getColumnHeaders(dynamicSheet, 3); // 4th row in Excel

            System.out.println("Static Headers: " + headerStaticSheet);
            System.out.println("Dynamic Headers: " + headerDynamicSheet);

            // Find common columns
            Set<String> commonColumns = new HashSet<>(headerStaticSheet);
            commonColumns.retainAll(headerDynamicSheet);

            // Get column indexes from correct rows
            Map<String, Integer> indexMapStatic = getColumnIndexMap(staticSheet, 0);
            Map<String, Integer> indexMapDynamic = getColumnIndexMap(dynamicSheet, 3);

            // Ensure flag column exists in Sheet 2
            if (!indexMapDynamic.containsKey(flagColumn)) {
                System.err.println("Error: Flag column '" + flagColumn + "' not found in Sheet 2.");
                return;
            }
            int flagColumnIndex = indexMapDynamic.get(flagColumn);

            // Process Sheet 2 rows for updating common columns
            for (int i = 4; i < dynamicSheet.getPhysicalNumberOfRows(); i++) { // Start from 5th row
                Row rowDynamic = dynamicSheet.getRow(i);
                Row rowStatic = staticSheet.getRow(i - 3); // Adjust row index for staticSheet

                if (rowDynamic == null || rowStatic == null) {
                    continue; // Skip empty rows
                }

                // Check if flag column is "Y"
                Cell flagCell = rowDynamic.getCell(flagColumnIndex);
                if (flagCell != null && "Y".equalsIgnoreCase(flagCell.getStringCellValue().trim())) {
                    // Update common column values
                    for (String column : commonColumns) {
                        int indexStatic = indexMapStatic.get(column);
                        int indexDynamic = indexMapDynamic.get(column);

                        Cell cellStatic = rowStatic.getCell(indexStatic);
                        Cell cellDynamic = rowDynamic.getCell(indexDynamic);

                        if (cellStatic != null) {
                            if (cellDynamic == null) {
                                cellDynamic = rowDynamic.createCell(indexDynamic);
                            }
                            cellDynamic.setCellValue(cellStatic.toString()); // Copy value from Sheet 1
                        }
                    }
                }
            }

            // Save updated data to output file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbookDynamic.write(fos);
            }

            System.out.println("Dynamic Sheet updated successfully!");

        } catch (IOException e) {
            System.err.println("Error processing the Excel file: " + e.getMessage());
        }
    }

    /**
     * Retrieves a map of column names to their index positions from the given header row.
     */
    private static Map<String, Integer> getColumnIndexMap(Sheet sheet, int headerRowNum) {
        Map<String, Integer> indexMap = new HashMap<>();
        Row headerRow = sheet.getRow(headerRowNum);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) { // Use getLastCellNum() instead of getPhysicalNumberOfCells()
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getCellType() == CellType.STRING) { // Check if cell is not null and is a string
                    String header = cell.getStringCellValue().trim();
                    if (!header.isEmpty()) { // Ignore empty headers
                        indexMap.put(header, i);
                    }
                }
            }
        }
        return indexMap;
    }

    /**
     * Retrieves a list of column headers from the specified row.
     */
    public static List<String> getColumnHeaders(Sheet sheet, int rowNum) {
        List<String> headers = new ArrayList<>();
        Row headerRow = sheet.getRow(rowNum);
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
