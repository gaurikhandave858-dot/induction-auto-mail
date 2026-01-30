import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Utility to clear data from master attendance sheet while preserving headers
 */
public class ClearMasterData {
    public static void main(String[] args) {
        String masterFilePath = "master_attendance_summary.xlsx";
        
        try {
            // Read the existing workbook
            FileInputStream fis = new FileInputStream(masterFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            
            // Get the last row number
            int lastRowNum = sheet.getLastRowNum();
            
            System.out.println("Found " + (lastRowNum) + " data rows to clear (excluding header row 0)");
            
            // Clear all rows except the header (row 0)
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
            
            // Alternatively, clear cell contents while keeping row structure
            // This ensures we preserve the formatting
            lastRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    int lastCellNum = row.getLastCellNum();
                    for (int j = 0; j < lastCellNum; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            cell.setCellValue(""); // Clear cell content
                        }
                    }
                }
            }
            
            fis.close();
            
            // Write the updated workbook back
            FileOutputStream fos = new FileOutputStream(masterFilePath);
            workbook.write(fos);
            workbook.close();
            fos.close();
            
            System.out.println("Master sheet data cleared successfully!");
            System.out.println("Headers preserved in row 1");
            System.out.println("Ready for new data update.");
            
        } catch (IOException e) {
            System.err.println("Error processing master file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}