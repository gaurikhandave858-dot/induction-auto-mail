import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Utility to check the structure of master attendance sheet
 */
public class CheckMasterStructure {
    public static void main(String[] args) {
        String masterFilePath = "master_attendance_summary.xlsx";
        
        try {
            FileInputStream fis = new FileInputStream(masterFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            
            System.out.println("=== Master Sheet Structure Analysis ===");
            System.out.println("Total rows in sheet: " + (sheet.getLastRowNum() + 1));
            
            // Check first 20 rows for headers
            System.out.println("\nFirst 20 rows content:");
            for (int i = 0; i <= Math.min(20, sheet.getLastRowNum()); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    System.out.print("Row " + i + ": ");
                    for (int j = 0; j < Math.min(10, row.getLastCellNum()); j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            String cellValue = getCellValue(cell);
                            if (!cellValue.trim().isEmpty()) {
                                System.out.print("[" + j + "]'" + cellValue + "' ");
                            }
                        }
                    }
                    System.out.println();
                }
            }
            
            // Find the actual header row
            System.out.println("\n=== Header Detection ===");
            for (int i = 0; i <= Math.min(30, sheet.getLastRowNum()); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    String rowContent = "";
                    boolean hasContent = false;
                    for (int j = 0; j < Math.min(15, row.getLastCellNum()); j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            String cellValue = getCellValue(cell);
                            if (!cellValue.trim().isEmpty()) {
                                rowContent += "'" + cellValue + "' ";
                                hasContent = true;
                            }
                        }
                    }
                    if (hasContent && rowContent.toLowerCase().contains("name") || 
                        rowContent.toLowerCase().contains("sr") || 
                        rowContent.toLowerCase().contains("no")) {
                        System.out.println("Potential header row " + i + ": " + rowContent);
                    }
                }
            }
            
            workbook.close();
            fis.close();
            
        } catch (IOException e) {
            System.err.println("Error reading master file: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}