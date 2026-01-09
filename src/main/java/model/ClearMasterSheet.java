package model;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ClearMasterSheet {
    public static void main(String[] args) {
        String masterFilePath = "master_attendance_summary.xlsx"; // Default name relative to where the program runs
        
        // If a specific file path is provided as argument
        if (args.length > 0) {
            masterFilePath = args[0];
        }
        
        clearMasterExcelFile(masterFilePath);
    }
    
    public static void clearMasterExcelFile(String filePath) {
        try {
            // Try to open existing file
            Workbook workbook;
            try (FileInputStream fis = new FileInputStream(filePath)) {
                if (filePath.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (filePath.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    System.out.println("Unsupported file format. Please use .xls or .xlsx files.");
                    return;
                }
            } catch (IOException e) {
                // File doesn't exist, create new workbook
                if (filePath.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook();
                } else {
                    workbook = new HSSFWorkbook();
                }
                System.out.println("Master file doesn't exist, creating a new one: " + filePath);
            }
            
            // Get or create the first sheet
            Sheet sheet = null;
            try {
                sheet = workbook.getSheetAt(0);
            } catch (IllegalArgumentException e) {
                // If there are no sheets, create one
                sheet = workbook.createSheet("AttendanceSummary");
            }
            if (sheet == null) {
                sheet = workbook.createSheet("AttendanceSummary");
            }
            
            // Clear all rows in the sheet
            int numberOfRows = sheet.getLastRowNum();
            for (int i = numberOfRows; i >= 0; i--) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
            
            // Create a fresh sheet to ensure it's completely empty
            String sheetName = sheet.getSheetName();
            workbook.removeSheetAt(0);
            workbook.createSheet(sheetName);
            
            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
            
            workbook.close();
            
            System.out.println("Master Excel file has been cleared: " + filePath);
            
        } catch (IOException e) {
            System.err.println("Error clearing master Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}