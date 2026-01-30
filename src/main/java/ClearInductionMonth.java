import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Utility to clear induction month data from master attendance sheet
 */
public class ClearInductionMonth {
    public static void main(String[] args) {
        String masterFilePath = "master_attendance_summary.xlsx";
        int inductionMonthColumn = 5; // Column E (0-based index = 4, but we'll use 1-based = 5)
        
        try {
            // Read the existing workbook
            FileInputStream fis = new FileInputStream(masterFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            
            // Get the last row number
            int lastRowNum = sheet.getLastRowNum();
            
            System.out.println("Found " + (lastRowNum) + " rows in master sheet");
            System.out.println("Clearing Induction Month column (Column " + inductionMonthColumn + ") data...");
            
            // Clear data in Induction Month column for all rows (starting from row 1, skipping header)
            int clearedCount = 0;
            for (int i = 1; i <= lastRowNum; i++) { // Start from row 1 (data rows)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(inductionMonthColumn - 1); // Convert to 0-based index
                    if (cell != null) {
                        // Clear the cell content
                        cell.setCellValue(""); // Set to empty string
                        clearedCount++;
                    }
                }
            }
            
            fis.close();
            
            // Write the updated workbook back
            FileOutputStream fos = new FileOutputStream(masterFilePath);
            workbook.write(fos);
            workbook.close();
            fos.close();
            
            System.out.println("Successfully cleared Induction Month data!");
            System.out.println("Cells cleared: " + clearedCount);
            System.out.println("Column " + inductionMonthColumn + " is now empty.");
            
        } catch (IOException e) {
            System.err.println("Error processing master file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}