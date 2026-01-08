package model;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProcessor {
    
    public List<Employee> readAttendanceData(String filePath) throws IOException {
        List<Employee> employees = new ArrayList<>();
        
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = null;
        
        try {
            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (filePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IOException("Unsupported file format. Please use .xls or .xlsx files.");
            }
            
            Sheet sheet = workbook.getSheetAt(0);
            
            // Find header row (first non-empty row)
            Row headerRow = null;
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null && !isRowEmpty(row)) {
                    headerRow = row;
                    break;
                }
            }
            
            if (headerRow == null) {
                throw new IOException("No header row found in Excel file");
            }
            
            // Identify column indices by header names
            int srNoCol = -1, nameCol = -1, contactNoCol = -1, day1Col = -1, day2Col = -1;
            String day1Date = "", day2Date = "";
            
            for (Cell cell : headerRow) {
                String header = getCellValueAsString(cell).trim();
                
                // Find Sr. No column - with more variations
                if (header.toLowerCase().contains("sr") || header.toLowerCase().contains("serial") || 
                    header.toLowerCase().contains("no") || header.equalsIgnoreCase("srno") ||
                    header.toLowerCase().contains("number") || header.toLowerCase().startsWith("sr")) {
                    srNoCol = cell.getColumnIndex();
                }
                // Find Name column - with more variations
                else if (header.toLowerCase().contains("name") || header.equalsIgnoreCase("employee") || 
                         header.equalsIgnoreCase("emp") || header.toLowerCase().startsWith("name")) {
                    nameCol = cell.getColumnIndex();
                }
                // Find Contact No column - with more variations
                else if (header.toLowerCase().contains("contact") || header.toLowerCase().contains("phone") || 
                         header.toLowerCase().contains("mobile") || header.toLowerCase().contains("tel") || 
                         header.toLowerCase().contains("number") || header.toLowerCase().contains("no")) {
                    contactNoCol = cell.getColumnIndex();
                }
                // Find Day 1 column by looking for date patterns
                else if (isDateColumn(header) && day1Col == -1) { // First date column found becomes Day 1
                    day1Col = cell.getColumnIndex();
                    day1Date = header; // Store the full header which contains the date
                }
                // Find Day 2 column by looking for date patterns
                else if (isDateColumn(header) && day1Col != -1 && day2Col == -1) { // Second date column found becomes Day 2
                    day2Col = cell.getColumnIndex();
                    day2Date = header; // Store the full header which contains the date
                }
            }
            
            // Validate that we found required columns
            if (srNoCol == -1 || nameCol == -1 || day1Col == -1 || day2Col == -1) {
                throw new IOException("Required columns not found. Excel file must contain: Sr. No, Name, and 2 date columns (e.g., 08/01/2026, 09/01/2026)");
            }
            
            // Process data rows
            for (int i = headerRow.getRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) {
                    continue; // Skip empty rows
                }
                
                Employee employee = new Employee();
                
                // Get Serial Number
                if (srNoCol < row.getLastCellNum()) {
                    Cell cell = row.getCell(srNoCol);
                    if (cell != null) {
                        try {
                            String cellValue = getCellValueAsString(cell);
                            int srNoValue = Integer.parseInt(cellValue);
                            employee.setSrNo(srNoValue);
                        } catch (NumberFormatException e) {
                            // If parsing fails, set to 0 or handle as needed
                            employee.setSrNo(0);
                        }
                    }
                }
                
                // Get Name
                if (nameCol < row.getLastCellNum()) {
                    Cell cell = row.getCell(nameCol);
                    if (cell != null) {
                        employee.setName(getCellValueAsString(cell));
                    }
                }
                
                // Get Contact No
                if (contactNoCol < row.getLastCellNum() && contactNoCol != -1) { // contact no is optional
                    Cell cell = row.getCell(contactNoCol);
                    if (cell != null) {
                        employee.setContactNo(getCellValueAsString(cell));
                    }
                }
                
                // Get Day 1 Attendance
                if (day1Col < row.getLastCellNum()) {
                    Cell cell = row.getCell(day1Col);
                    if (cell != null) {
                        employee.setDay1Attendance(getCellValueAsString(cell));
                        employee.setDay1Date(day1Date); // Set the date information
                    }
                }
                
                // Get Day 2 Attendance
                if (day2Col < row.getLastCellNum()) {
                    Cell cell = row.getCell(day2Col);
                    if (cell != null) {
                        employee.setDay2Attendance(getCellValueAsString(cell));
                        employee.setDay2Date(day2Date); // Set the date information
                    }
                }
                
                employees.add(employee);
            }
            
            workbook.close();
        } finally {
            fis.close();
        }
        
        return employees;
    }
    
    /**
     * Updates the master Excel sheet with attendance records
     */
    public void updateMasterSheet(List<EmailRecord> emailRecords, String masterFilePath) throws IOException {
        Workbook workbook;
        
        // Try to open existing file, if it doesn't exist create new one
        try (FileInputStream fis = new FileInputStream(masterFilePath)) {
            if (masterFilePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (masterFilePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                workbook = new XSSFWorkbook(); // Default to xlsx
            }
        } catch (IOException e) {
            // File doesn't exist, create new workbook
            workbook = new XSSFWorkbook();
        }
        
        Sheet sheet = workbook.getSheet("AttendanceSummary");
        if (sheet == null) {
            sheet = workbook.createSheet("AttendanceSummary");
        }
        
        // Create header row if it's a new sheet
        if (sheet.getLastRowNum() == 0) {
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Record Info");
            headerRow.createCell(1).setCellValue("Value");
        }
        
        // Add new records in vertical format
        int lastRowNum = sheet.getLastRowNum();
        for (EmailRecord record : emailRecords) {
            // Row 1: Date
            Row dateRow = sheet.createRow(lastRowNum + 1);
            dateRow.createCell(0).setCellValue("Date");
            dateRow.createCell(1).setCellValue(record.getDate());
            
            // Row 2: Day 1 Present Count
            Row row1 = sheet.createRow(lastRowNum + 2);
            row1.createCell(0).setCellValue("Day 1 Present Count");
            row1.createCell(1).setCellValue(record.getDay1PresentCount());
            
            // Row 3: Day 1 Absent Count
            Row row2 = sheet.createRow(lastRowNum + 3);
            row2.createCell(0).setCellValue("Day 1 Absent Count");
            row2.createCell(1).setCellValue(record.getDay1AbsentCount());
            
            // Row 4: Day 1 Absent Names
            Row row3 = sheet.createRow(lastRowNum + 4);
            row3.createCell(0).setCellValue("Day 1 Absent Names");
            row3.createCell(1).setCellValue(record.getDay1AbsentNames());
            
            // Row 5: Day 2 Present Count
            Row row4 = sheet.createRow(lastRowNum + 5);
            row4.createCell(0).setCellValue("Day 2 Present Count");
            row4.createCell(1).setCellValue(record.getDay2PresentCount());
            
            // Row 6: Day 2 Absent Count
            Row row5 = sheet.createRow(lastRowNum + 6);
            row5.createCell(0).setCellValue("Day 2 Absent Count");
            row5.createCell(1).setCellValue(record.getDay2AbsentCount());
            
            // Row 7: Day 2 Absent Names
            Row row6 = sheet.createRow(lastRowNum + 7);
            row6.createCell(0).setCellValue("Day 2 Absent Names");
            row6.createCell(1).setCellValue(record.getDay2AbsentNames());
            
            // Row 8: Report Generated On date
            Row row7 = sheet.createRow(lastRowNum + 8);
            row7.createCell(0).setCellValue("Report Generated On date");
            row7.createCell(1).setCellValue(java.time.LocalDate.now().toString());
            
            // Row 9: Complete Email Content (formatted like the email summary)
            String emailContent = "EMAIL SUMMARY CONTENT\n\n" +
                                 "Report Generated On: " + java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")) + "\n\n" +
                                 "Day 1:\n" +
                                 "Day 1 (" + (record.getDay1DateString() != null ? record.getDay1DateString() : "") + ") Summary:\n" +
                                 "Present: " + record.getDay1PresentCount() + "\n" +
                                 "Absent: " + record.getDay1AbsentCount() + "\n\n" +
                                 "Day 1 (" + (record.getDay1DateString() != null ? record.getDay1DateString() : "") + ") Absent Employee List:\n\n";
            
            if (!record.getDay1AbsentNames().isEmpty()) {
                String[] day1Names = record.getDay1AbsentNames().split(", ");
                for (String name : day1Names) {
                    if (!name.trim().isEmpty()) {
                        emailContent += "- " + name.trim() + "\n";
                    }
                }
            } else {
                emailContent += "- None\n";
            }
            
            emailContent += "\n" +
                             "Day 2 (" + (record.getDay2DateString() != null ? record.getDay2DateString() : "") + ") Summary:\n" +
                             "Present: " + record.getDay2PresentCount() + "\n" +
                             "Absent: " + record.getDay2AbsentCount() + "\n\n" +
                             "Day 2 (" + (record.getDay2DateString() != null ? record.getDay2DateString() : "") + ") Absent Employee List:\n\n";
            
            if (!record.getDay2AbsentNames().isEmpty()) {
                String[] day2Names = record.getDay2AbsentNames().split(", ");
                for (String name : day2Names) {
                    if (!name.trim().isEmpty()) {
                        emailContent += "- " + name.trim() + "\n";
                    }
                }
            } else {
                emailContent += "- None\n";
            }
            
            Row row8 = sheet.createRow(lastRowNum + 9);
            row8.createCell(0).setCellValue("Complete Email Content");
            row8.createCell(1).setCellValue(emailContent);
            
            // Add a blank row for spacing between records
            Row spacingRow = sheet.createRow(lastRowNum + 10);
            spacingRow.createCell(0).setCellValue("");
            spacingRow.createCell(1).setCellValue("");
            
            lastRowNum += 10; // Move to next position (9 data rows + 1 spacing row)
        }
        
        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(masterFilePath)) {
            workbook.write(fileOut);
        }
        
        workbook.close();
    }
    
    /**
     * Gets cell value as string regardless of cell type
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                // Check if it's a date
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // For numeric values, return as string
                    double value = cell.getNumericCellValue();
                    // If it's a whole number, return as integer string
                    if (value == Math.floor(value)) {
                        return String.valueOf((int) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    // Evaluate the formula to get the resulting value
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case STRING:
                            return cellValue.getStringValue().trim();
                        case NUMERIC:
                            double value = cellValue.getNumberValue();
                            if (value == Math.floor(value)) {
                                return String.valueOf((int) value);
                            } else {
                                return String.valueOf(value);
                            }
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    // If evaluation fails, return the formula as string
                    return cell.getCellFormula();
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }
    
    /**
     * Checks if a row is empty (all cells are blank)
     */
    private boolean isRowEmpty(Row row) {
        if (row == null) return true;
        
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
    
    /**
     * Checks if a column header represents a date
     * Looks for common date patterns in the header
     */
    private boolean isDateColumn(String header) {
        if (header == null || header.trim().isEmpty()) {
            return false;
        }
        
        String lowerHeader = header.toLowerCase().trim();
        
        // Look for common date patterns
        // DD/MM/YYYY or DD-MM-YYYY or DD/MM/YY or DD-MM-YY
        if (lowerHeader.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
            return true;
        }
        
        // Handle case where the header might have extra spaces or formatting
        String trimmedHeader = header.replaceAll("\\s+", ""); // Remove all whitespace
        if (trimmedHeader.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
            return true;
        }
        
        // Also check the original header without converting to lowercase for date patterns
        String origHeader = header.trim();
        if (origHeader.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
            return true;
        }
        
        // Look for date-like patterns in various formats
        // MM/DD/YYYY, DD/MM/YYYY, etc.
        if (lowerHeader.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
            return true;
        }
        
        // Additional check for date in parentheses format
        if (origHeader.contains("(") && origHeader.contains(")")) {
            int start = origHeader.indexOf('(');
            int end = origHeader.indexOf(')');
            String insideBrackets = origHeader.substring(start + 1, end);
            if (insideBrackets.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
                return true;
            }
        }
        
        // Check if it contains date-like keywords with actual dates
        if (lowerHeader.contains("/") || lowerHeader.contains("-")) {
            // Split by separator and check if parts look like date components
            String[] parts = lowerHeader.split("[\\\\/-]");
            if (parts.length == 3) {
                try {
                    int part1 = Integer.parseInt(parts[0].trim());
                    int part2 = Integer.parseInt(parts[1].trim());
                    int part3 = Integer.parseInt(parts[2].trim());
                    
                    // Basic validation: day (1-31), month (1-12), year (reasonable range)
                    boolean isValidDay = (part1 >= 1 && part1 <= 31);
                    boolean isValidMonth = (part2 >= 1 && part2 <= 12) || (part1 >= 1 && part1 <= 12);
                    boolean isValidYear = (part3 >= 2000 && part3 <= 2100) || (part3 >= 0 && part3 <= 99);
                    
                    return (isValidDay && isValidMonth && isValidYear) || 
                           (isValidMonth && isValidDay && isValidYear);
                } catch (NumberFormatException e) {
                    // Not a valid number combination
                }
            }
        }
        
        // Check for date-like formats in brackets or parentheses
        if (lowerHeader.contains("(") && lowerHeader.contains(")")) {
            int start = lowerHeader.indexOf('(');
            int end = lowerHeader.indexOf(')');
            String insideBrackets = lowerHeader.substring(start + 1, end);
            
            if (insideBrackets.matches("\\d{1,2}[\\\\/-]\\d{1,2}[\\\\/-]\\d{2,4}")) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Normalizes header text by converting to lowercase and removing special characters
     */
    private String normalizeHeader(String header) {
        if (header == null) return "";
        
        // Convert to lowercase and remove special characters except alphanumeric and spaces
        return header.toLowerCase().replaceAll("[^a-z0-9\\s]", " ").trim();
    }
    
    /**
     * Creates a formatted Excel sheet with attendance summary in the exact format required
     */
    public void createFormattedAttendanceSummarySheet(AttendanceSummary summary, String day1Date, String day2Date, String fileName) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Attendance Summary");
        
        // Create font styles
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        
        // Create cell styles
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        
        int rowNum = 0;
        
        // EMAIL SUMMARY CONTENT
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(0);
        cell.setCellValue("EMAIL SUMMARY CONTENT");
        cell.setCellStyle(boldStyle);
        
        // Report Generated On line
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Report Generated On: " + java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        
        // Blank row
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 1 section
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Day 1 (" + day1Date + ")");
        cell.setCellStyle(boldStyle);
        
        // Day 1 Present Count
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Present Count: " + summary.getDay1PresentCount());
        
        // Day 1 Absent Count
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Absent Count: " + summary.getDay1AbsentCount());
        
        // Day 1 Absent Employee List header
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Absent Employee List:");
        
        // Day 1 Absent Employee List
        if (!summary.getDay1AbsentNames().isEmpty()) {
            String[] day1Names = summary.getDay1AbsentNames().split(", ");
            for (String name : day1Names) {
                if (!name.trim().isEmpty()) {
                    row = sheet.createRow(rowNum++);
                    cell = row.createCell(0);
                    cell.setCellValue("- " + name.trim());
                }
            }
        } else {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue("- None");
        }
        
        // Blank row
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 2 section
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Day 2 (" + day2Date + ")");
        cell.setCellStyle(boldStyle);
        
        // Day 2 Present Count
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Present Count: " + summary.getDay2PresentCount());
        
        // Day 2 Absent Count
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Absent Count: " + summary.getDay2AbsentCount());
        
        // Day 2 Absent Employee List header
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("Absent Employee List:");
        
        // Day 2 Absent Employee List
        if (!summary.getDay2AbsentNames().isEmpty()) {
            String[] day2Names = summary.getDay2AbsentNames().split(", ");
            for (String name : day2Names) {
                if (!name.trim().isEmpty()) {
                    row = sheet.createRow(rowNum++);
                    cell = row.createCell(0);
                    cell.setCellValue("- " + name.trim());
                }
            }
        } else {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue("- None");
        }
        
        // Auto-size the column to make it readable
        sheet.autoSizeColumn(0);
        
        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }
        
        workbook.close();
    }
    
    /**
     * Inserts email summary content into an existing Excel sheet starting from the next available row
     */
    public void insertEmailSummaryIntoExistingSheet(EmailRecord record, String masterFilePath, int startRow) throws IOException {
        Workbook workbook;
        
        // Open existing file
        try (FileInputStream fis = new FileInputStream(masterFilePath)) {
            if (masterFilePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (masterFilePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IOException("Unsupported file format. Please use .xls or .xlsx files.");
            }
        }
        
        Sheet sheet = workbook.getSheetAt(0); // Use first sheet
        
        // Find the next available row to append data (after any existing content)
        int lastRowNum = sheet.getLastRowNum();
        int currentRow = Math.max(lastRowNum + 1, startRow); // Use either the next empty row or startRow, whichever is greater
        
        // Create font styles for bold text
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        
        // Create cell style for bold text
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        
        // Add a blank row before the new summary for separation
        if (lastRowNum >= 0) { // If there's existing content
            currentRow++;
        }
        
        // EMAIL SUMMARY CONTENT (bold)
        Row row = sheet.createRow(currentRow++);
        Cell cell = row.createCell(0);
        cell.setCellValue("EMAIL SUMMARY CONTENT");
        cell.setCellStyle(boldStyle);
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Report Generated On line
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Report Generated On: " + java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 1 header
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Day 1:");
        
        // Day 1 section with date and summary
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Day 1 (" + (record.getDay1DateString() != null ? record.getDay1DateString() : "") + ") Summary:");
        
        // Day 1 Present Count
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Present: " + record.getDay1PresentCount());
        
        // Day 1 Absent Count
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Absent: " + record.getDay1AbsentCount());
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 1 Absent Employee List header
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Day 1 (" + (record.getDay1DateString() != null ? record.getDay1DateString() : "") + ") Absent Employee List:");
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 1 Absent Employee List
        if (!record.getDay1AbsentNames().isEmpty()) {
            String[] day1Names = record.getDay1AbsentNames().split(", ");
            for (String name : day1Names) {
                if (!name.trim().isEmpty()) {
                    row = sheet.createRow(currentRow++);
                    cell = row.createCell(0);
                    cell.setCellValue("- " + name.trim());
                }
            }
        } else {
            row = sheet.createRow(currentRow++);
            cell = row.createCell(0);
            cell.setCellValue("- None");
        }
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 2 section with date and summary
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Day 2 (" + (record.getDay2DateString() != null ? record.getDay2DateString() : "") + ") Summary:");
        
        // Day 2 Present Count
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Present: " + record.getDay2PresentCount());
        
        // Day 2 Absent Count
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Absent: " + record.getDay2AbsentCount());
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 2 Absent Employee List header
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("Day 2 (" + (record.getDay2DateString() != null ? record.getDay2DateString() : "") + ") Absent Employee List:");
        
        // Blank row
        row = sheet.createRow(currentRow++);
        cell = row.createCell(0);
        cell.setCellValue("");
        
        // Day 2 Absent Employee List
        if (!record.getDay2AbsentNames().isEmpty()) {
            String[] day2Names = record.getDay2AbsentNames().split(", ");
            for (String name : day2Names) {
                if (!name.trim().isEmpty()) {
                    row = sheet.createRow(currentRow++);
                    cell = row.createCell(0);
                    cell.setCellValue("- " + name.trim());
                }
            }
        } else {
            row = sheet.createRow(currentRow++);
            cell = row.createCell(0);
            cell.setCellValue("- None");
        }
        
        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(masterFilePath)) {
            workbook.write(fileOut);
        }
        
        workbook.close();
    }
}