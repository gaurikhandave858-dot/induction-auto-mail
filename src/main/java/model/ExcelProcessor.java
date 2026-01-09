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
            
            // Find header row (first non-empty row with actual data columns)
            Row headerRow = null;
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null && !isRowEmpty(row)) {
                    // Look for the row that contains the actual column headers (Sr No, Name, etc.)
                    // This should be after any training/module description rows
                    if (hasDataColumns(row)) {
                        headerRow = row;
                        break;
                    }
                }
            }
            
            if (headerRow == null) {
                throw new IOException("No header row found in Excel file");
            }
            
            // Identify column indices by header names
            int srNoCol = -1, pNoCol = -1, nameCol = -1, genderCol = -1, tradeCol = -1, mobileNoCol = -1;
            int preTestCol = -1, postTestCol = -1, reExamCol = -1, deployShopCol = -1;
            
            // List to hold all date columns (attendance columns)
            List<Integer> dateCols = new ArrayList<>();
            List<String> dateHeaders = new ArrayList<>();
            
            for (Cell cell : headerRow) {
                String header = getCellValueAsString(cell).trim().toLowerCase();
                
                // Find various column types
                if (header.contains("sr") && (header.contains("no") || header.contains("num"))) {
                    srNoCol = cell.getColumnIndex();
                } else if (header.contains("p") && header.contains("no")) {
                    pNoCol = cell.getColumnIndex();
                } else if (header.contains("name") && !header.contains("post") && !header.contains("pre")) {
                    nameCol = cell.getColumnIndex();
                } else if (header.contains("gender")) {
                    genderCol = cell.getColumnIndex();
                } else if (header.contains("trade")) {
                    tradeCol = cell.getColumnIndex();
                } else if (header.contains("mobile") || header.contains("contact") || (header.contains("no") && !header.contains("sr") && !header.contains("p"))) {
                    mobileNoCol = cell.getColumnIndex();
                } else if (header.contains("pre") && header.contains("test")) {
                    preTestCol = cell.getColumnIndex();
                } else if (header.contains("post") && header.contains("test")) {
                    postTestCol = cell.getColumnIndex();
                } else if (header.contains("re") && header.contains("exam")) {
                    reExamCol = cell.getColumnIndex();
                } else if (header.contains("deploy") && header.contains("shop")) {
                    deployShopCol = cell.getColumnIndex();
                } else if (isDateColumn(header)) {
                    // This is a date column, add to our list of date columns
                    dateCols.add(cell.getColumnIndex());
                    dateHeaders.add(getCellValueAsString(cell).trim()); // Store the original header
                }
            }
            
            // Validate that we found required columns (at least Sr No, Name, and 2 date columns)
            if (srNoCol == -1 || nameCol == -1 || dateCols.size() < 2) {
                throw new IOException("Required columns not found. Excel file must contain: Sr No, Name, and at least 2 date columns (e.g., 16-Dec, 17-Dec)");
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
                
                // Get P No (if exists)
                if (pNoCol < row.getLastCellNum() && pNoCol != -1) {
                    Cell cell = row.getCell(pNoCol);
                    if (cell != null) {
                        employee.setPNo(getCellValueAsString(cell));
                    }
                }
                
                // Get Name
                if (nameCol < row.getLastCellNum()) {
                    Cell cell = row.getCell(nameCol);
                    if (cell != null) {
                        employee.setName(getCellValueAsString(cell));
                    }
                }
                
                // Get Gender (if exists)
                if (genderCol < row.getLastCellNum() && genderCol != -1) {
                    Cell cell = row.getCell(genderCol);
                    if (cell != null) {
                        employee.setGender(getCellValueAsString(cell));
                    }
                }
                
                // Get Trade (if exists)
                if (tradeCol < row.getLastCellNum() && tradeCol != -1) {
                    Cell cell = row.getCell(tradeCol);
                    if (cell != null) {
                        employee.setTrade(getCellValueAsString(cell));
                    }
                }
                
                // Get Mobile No (if exists)
                if (mobileNoCol < row.getLastCellNum() && mobileNoCol != -1) {
                    Cell cell = row.getCell(mobileNoCol);
                    if (cell != null) {
                        employee.setMobileNo(getCellValueAsString(cell));
                    }
                }
                
                // Get Pre-Test (if exists)
                if (preTestCol < row.getLastCellNum() && preTestCol != -1) {
                    Cell cell = row.getCell(preTestCol);
                    if (cell != null) {
                        employee.setPreTest(getCellValueAsString(cell));
                    }
                }
                
                // Get Post-Test (if exists)
                if (postTestCol < row.getLastCellNum() && postTestCol != -1) {
                    Cell cell = row.getCell(postTestCol);
                    if (cell != null) {
                        employee.setPostTest(getCellValueAsString(cell));
                    }
                }
                
                // Get Re-Exam (if exists)
                if (reExamCol < row.getLastCellNum() && reExamCol != -1) {
                    Cell cell = row.getCell(reExamCol);
                    if (cell != null) {
                        employee.setReExam(getCellValueAsString(cell));
                    }
                }
                
                // Get Deploy Shop (if exists)
                if (deployShopCol < row.getLastCellNum() && deployShopCol != -1) {
                    Cell cell = row.getCell(deployShopCol);
                    if (cell != null) {
                        employee.setDeployShop(getCellValueAsString(cell));
                    }
                }
                
                // Process attendance for each date column
                if (dateCols.size() > 0) {
                    // For backward compatibility, assign first two date columns to Day1 and Day2
                    if (dateCols.size() > 0) {
                        Cell cell = row.getCell(dateCols.get(0));
                        if (cell != null) {
                            employee.setDay1Attendance(getCellValueAsString(cell));
                            employee.setDay1Date(dateHeaders.get(0));
                        }
                    }
                    
                    if (dateCols.size() > 1) {
                        Cell cell = row.getCell(dateCols.get(1));
                        if (cell != null) {
                            employee.setDay2Attendance(getCellValueAsString(cell));
                            employee.setDay2Date(dateHeaders.get(1));
                        }
                    }
                    
                    // Store all date attendances for future use
                    employee.setAllDateAttendances(new ArrayList<>());
                    employee.setAllDateHeaders(new ArrayList<>());
                    
                    for (int j = 0; j < dateCols.size(); j++) {
                        Cell cell = row.getCell(dateCols.get(j));
                        if (cell != null) {
                            employee.getAllDateAttendances().add(getCellValueAsString(cell));
                            employee.getAllDateHeaders().add(dateHeaders.get(j));
                        }
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
     * Checks if a row contains data columns (like Sr No, Name, etc.)
     * Used to skip training/module description rows at the top of the Excel sheet
     */
    private boolean hasDataColumns(Row row) {
        if (row == null) return false;
        
        for (Cell cell : row) {
            String header = getCellValueAsString(cell).toLowerCase().trim();
            
            // Check for common data column headers
            if (header.contains("sr") && (header.contains("no") || header.contains("num")) ||
                header.contains("p") && header.contains("no") ||
                header.contains("name") ||
                header.contains("gender") ||
                header.contains("trade") ||
                header.contains("mobile") ||
                header.contains("contact") ||
                header.contains("no") && !header.contains("module") ||
                header.contains("pre") && header.contains("test") ||
                header.contains("post") && header.contains("test") ||
                header.contains("re") && header.contains("exam") ||
                header.contains("deploy") && header.contains("shop") ||
                isDateColumn(header)) {
                return true;
            }
        }
        return false;
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
        
        // Check for date-like formats in brackets or parentheses
        if (lowerHeader.contains("(") && lowerHeader.contains(")")) {
            int start = lowerHeader.indexOf('(');
            int end = lowerHeader.indexOf(')');
            String insideBrackets = lowerHeader.substring(start + 1, end);
            
            if (insideBrackets.matches("\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4}")) {
                return true;
            }
        }
        
        // Additional check for date in parentheses format
        if (origHeader.contains("(") && origHeader.contains(")")) {
            int start = origHeader.indexOf('(');
            int end = origHeader.indexOf(')');
            String insideBrackets = origHeader.substring(start + 1, end);
            if (insideBrackets.matches("\\\\d{1,2}[\\\\/-]\\\\d{1,2}[\\\\/-]\\\\d{2,4}")) {
                return true;
            }
        }
        
        // Check for date-like formats in brackets or parentheses
        if (lowerHeader.contains("(") && lowerHeader.contains(")")) {
            int start = lowerHeader.indexOf('(');
            int end = lowerHeader.indexOf(')');
            String insideBrackets = lowerHeader.substring(start + 1, end);
            
            if (insideBrackets.matches("\\\\d{1,2}[\\\\/-]\\\\d{1,2}[\\\\/-]\\\\d{2,4}")) {
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
        
        // Additional check for dd-MMM format (e.g., 16-Dec, 17-Dec) as specified in SRS
        if (lowerHeader.matches("\\d{1,2}-[a-z]{3}")) {
            return true;
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
    
    public void insertEmailSummaryIntoExistingSheet(EmailRecord record, String masterFilePath, int startRowNum) throws IOException {
        FileInputStream fis = new FileInputStream(masterFilePath);
        Workbook workbook;
        
        // Determine file type and create appropriate workbook
        if (masterFilePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(fis);
        } else if (masterFilePath.endsWith(".xls")) {
            workbook = new HSSFWorkbook(fis);
        } else {
            fis.close();
            throw new IOException("Unsupported file format. Please use .xls or .xlsx files.");
        }
        
        fis.close();
        
        Sheet sheet = workbook.getSheetAt(0);
        
        // If sheet doesn't exist, create one
        if (sheet == null) {
            sheet = workbook.createSheet("AttendanceSummary");
        }
        
        // Get or create a bold font style
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        
        // Find the last row to add new content
        int currentRow = startRowNum;
        
        // Add a blank row before the new summary for separation
        if (sheet.getLastRowNum() >= 0) { // If there's existing content
            currentRow = sheet.getLastRowNum() + 2; // Leave one blank row for separation
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
    
    public void updateMasterSheet(List<EmailRecord> records, String fileName) throws IOException {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("AttendanceSummary");
        
        // Create font styles
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);
        
        // Add header row
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Master Attendance Summary");
        headerCell.setCellStyle(boldStyle);
        
        // Add initial data if any records exist
        int rowNum = 2; // Start after header and blank row
        for (EmailRecord record : records) {
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(0);
            cell.setCellValue("Date: " + record.getDate());
            
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue("Day 1 - Present: " + record.getDay1PresentCount() + ", Absent: " + record.getDay1AbsentCount());
            
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue("Day 2 - Present: " + record.getDay2PresentCount() + ", Absent: " + record.getDay2AbsentCount());
            
            rowNum++; // Blank row
        }
        
        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }
        
        workbook.close();
        
        System.out.println("Master sheet updated successfully: " + fileName);
    }
    
}