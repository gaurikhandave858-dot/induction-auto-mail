import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ClearMasterSheet {
    public static void main(String[] args) {
        String masterFilePath = "master_attendance_summary.xlsx";
        
        try {
            // Try to load existing workbook
            Workbook workbook;
            try {
                FileInputStream fis = new FileInputStream(masterFilePath);
                if (masterFilePath.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (masterFilePath.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    fis.close();
                    throw new IOException("Unsupported file format. Please use .xls or .xlsx files.");
                }
                fis.close();
            } catch (java.io.FileNotFoundException e) {
                // If file doesn't exist, create a new workbook
                if (masterFilePath.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook();
                } else {
                    workbook = new HSSFWorkbook();
                }
            }
            
            // Get or create the sheet - check if any sheets exist
            Sheet sheet;
            if (workbook.getNumberOfSheets() > 0) {
                // If there are existing sheets, get the first one
                sheet = workbook.getSheetAt(0);
            } else {
                // If no sheets exist, create one
                sheet = workbook.createSheet("AttendanceSummary");
            }
            
            // Get or create a bold font style for headers
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);
            
            // Clear all data rows except the header row
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= lastRowNum; i++) {  // Start from 1 to keep header row
                Row row = sheet.getRow(i);
                if (row != null) {
                    // Instead of removing the row, let's clear its cells
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            cell.setCellValue("");
                        }
                    }
                }
            }
            
            // Ensure the header row exists with the required 18 columns (updated structure)
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                headerRow = sheet.createRow(0);
                
                // Set up the master sheet headers according to requirements
                String[] headers = {
                    "Sr No", "P No", "Name", "Gender", "Induction Month", "Induction Week", 
                    "Induction Start Date", "Induction End Date", "Induction Day 1", 
                    "Induction Day 2", "Induction Day 3", "Induction Total HRS", 
                    "Pre-Test", "Induction Marks", "Induction Re-Exam", 
                    "Induction Total", "Exam Status", "Shop"
                };
                
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(boldStyle);
                }
            } else {
                // If header row exists, make sure it has the correct headers
                String[] headers = {
                    "Sr No", "P No", "Name", "Gender", "Induction Month", "Induction Week", 
                    "Induction Start Date", "Induction End Date", "Induction Day 1", 
                    "Induction Day 2", "Induction Day 3", "Induction Total HRS", 
                    "Pre-Test", "Induction Marks", "Induction Re-Exam", 
                    "Induction Total", "Exam Status", "Shop"
                };
                
                // Clear existing header cells and set new ones
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        cell.setCellValue("");
                    }
                }
                
                // Set up the master sheet headers according to requirements
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell == null) {
                        cell = headerRow.createCell(i);
                    }
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(boldStyle);
                }
            }
            
            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream(masterFilePath)) {
                workbook.write(fileOut);
            }
            
            workbook.close();
            
            System.out.println("Master sheet cleared successfully. Headers preserved.");
            
        } catch (IOException e) {
            System.err.println("Error clearing master sheet: " + e.getMessage());
            e.printStackTrace();
        }
    }
}