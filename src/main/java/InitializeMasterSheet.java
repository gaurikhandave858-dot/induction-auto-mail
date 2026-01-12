import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import model.ExcelProcessor;

public class InitializeMasterSheet {        
    public static void main(String[] args) {
        try {
            ExcelProcessor excelProcessor = new ExcelProcessor();                       
            
            // Create the master sheet with the proper headers
            createMasterSheetWithHeaders("master_attendance_summary.xlsx");
            
            System.out.println("Master attendance summary Excel sheet created successfully!");                                                  
            System.out.println("The sheet is located at: master_attendance_summary.xlsx");                                                      
            System.out.println("It will store all summary emails with both historical and recent data.");                                   
        } catch (Exception e) {
            System.err.println("Error creating master sheet: " + e.getMessage());                   
            e.printStackTrace();
        }
    }
    
    private static void createMasterSheetWithHeaders(String fileName) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("AttendanceSummary");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            
            // Define headers according to requirements
            String[] headers = {
                "Sr No", "P No", "Name", "Gender", "Batch No", "Grade", 
                "Induction Month", "Induction Start Date", "Induction End Date", 
                "Induction Day 1", "Induction Day 2", "Induction Day 3", 
                "Induction Total HRS", "Pre-Test", "Induction Marks", 
                "Induction Re-Exam", "Induction Total", "Exam Status", "Shop"
            };
            
            // Create bold font for headers
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);
            
            // Add headers to the row
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(boldStyle);
            }
            
            // Write the workbook to file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            System.err.println("Error creating master sheet: " + e.getMessage());
            e.printStackTrace();
        }
    }
}