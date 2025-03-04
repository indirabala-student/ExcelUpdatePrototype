package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        long startTime = System.nanoTime(); // Start Timer

        System.out.println("Starting Excel Processing...");

        String staticFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\RCPLNP01.xlsx";
        String dynamicFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\example.xlsx";
        String outputFilePath = "C:\\Users\\ibattula\\IdeaProjects\\TyageshProject\\src\\main\\java\\org\\example\\out_data.xlsx";
        String flagColumn = "Flag";

        try (Workbook workbookStatic = new XSSFWorkbook(new FileInputStream(staticFilePath));
             Workbook workbookDynamic = new XSSFWorkbook(new FileInputStream(dynamicFilePath))) {

            Sheet staticSheet = workbookStatic.getSheetAt(0);
            Sheet dynamicSheet = workbookDynamic.getSheet("000A");

            // Get column index mappings directly
            Map<String, Integer> indexMapStatic = getColumnIndexMap(staticSheet, 0);
            Map<String, Integer> indexMapDynamic = getColumnIndexMap(dynamicSheet, 3);

            // Identify common columns
            Set<String> commonColumns = new HashSet<>(indexMapStatic.keySet());
            commonColumns.retainAll(indexMapDynamic.keySet());

            // Ensure flag column exists in Sheet 2
            Integer flagColumnIndex = indexMapDynamic.get(flagColumn);
            if (flagColumnIndex == null) {
                System.err.println("Error: Flag column '" + flagColumn + "' not found in Sheet 2.");
                return;
            }

            // Process rows from Sheet 2
            for (int i = 4; i <= dynamicSheet.getLastRowNum(); i++) { // Start from 5th row
                Row rowDynamic = dynamicSheet.getRow(i);
                Row rowStatic = staticSheet.getRow(i - 3); // Adjust row index for staticSheet

                if (rowDynamic == null || rowStatic == null) continue;

                // Check flag column
                Cell flagCell = rowDynamic.getCell(flagColumnIndex);
                if (flagCell == null || !"Y".equalsIgnoreCase(flagCell.getStringCellValue().trim())) continue;

                // Update common columns
                for (String column : commonColumns) {
                    Integer indexStatic = indexMapStatic.get(column);
                    Integer indexDynamic = indexMapDynamic.get(column);

                    if (indexStatic == null || indexDynamic == null) continue;

                    Cell cellStatic = rowStatic.getCell(indexStatic);
                    if (cellStatic == null) continue;

                    Cell cellDynamic = rowDynamic.getCell(indexDynamic, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellDynamic.setCellValue(cellStatic.toString());
                }
            }

            // Save output
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbookDynamic.write(fos);
            }

            System.out.println("Dynamic Sheet updated successfully!");

        } catch (IOException e) {
            System.err.println("Error processing the Excel file: " + e.getMessage());
        }

        long endTime = System.nanoTime(); // End Timer
        System.out.println("Execution Time: " + (endTime - startTime) / 1e6 + " ms");
    }

    /**
     * Returns a map of column names to their index positions.
     */
    private static Map<String, Integer> getColumnIndexMap(Sheet sheet, int headerRowNum) {
        Map<String, Integer> indexMap = new HashMap<>();
        Row headerRow = sheet.getRow(headerRowNum);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String header = cell.getStringCellValue().trim();
                    if (!header.isEmpty()) indexMap.put(header, i);
                }
            }
        }
        return indexMap;
    }
}
