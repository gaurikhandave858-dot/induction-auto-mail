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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProcessor {

    private DataFormatter dataFormatter = new DataFormatter();
    private List<String> lastHeaders = new ArrayList<>();

    public List<String> getLastHeaders() {
        return lastHeaders;
    }

    public List<Employee> readAttendanceData(String filePath) throws IOException {
        List<Employee> employees = new ArrayList<>();
        this.lastHeaders = new ArrayList<>(); // Reset headers

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

            // If we couldn't find a header row using the standard method, try a more
            // lenient approach
            if (headerRow == null) {
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && !isRowEmpty(row)) {
                        // Look for the first row that contains "Sr No" and "Name" or date-like columns
                        if (containsKeyHeaders(row)) {
                            headerRow = row;
                            break;
                        }
                    }
                }
            }

            // If still no header found, try a more general approach looking for any row
            // with recognizable column headers
            if (headerRow == null) {
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && !isRowEmpty(row)) {
                        // Count how many cells in this row might be headers
                        int possibleHeaders = 0;
                        for (Cell cell : row) {
                            String cellValue = getCellValueAsString(cell).toLowerCase().trim();
                            if (cellValue.contains("name") || isDateColumn(cellValue) ||
                                    cellValue.contains("sr") || cellValue.contains("no") ||
                                    cellValue.contains("p ") || cellValue.contains("mobile") ||
                                    cellValue.contains("gender") || cellValue.contains("trade")) {
                                possibleHeaders++;
                            }
                        }

                        // If at least 2 cells look like possible headers, consider this the header row
                        if (possibleHeaders >= 2) {
                            headerRow = row;
                            break;
                        }
                    }
                }
            }

            // If still no header found, try to find a row that has both 'Name' and a date
            // column
            if (headerRow == null) {
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && !isRowEmpty(row)) {
                        boolean hasName = false;
                        boolean hasDate = false;

                        for (Cell cell : row) {
                            String cellValue = getCellValueAsString(cell).toLowerCase().trim();
                            if (cellValue.contains("name")) {
                                hasName = true;
                            }
                            if (isDateColumn(cellValue)) {
                                hasDate = true;
                            }
                        }

                        if (hasName && hasDate) {
                            headerRow = row;
                            break;
                        }
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
                } else if (header.contains("mobile") || header.contains("contact")
                        || (header.contains("no") && !header.contains("sr") && !header.contains("p"))) {
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
                    dateCols.add(cell.getColumnIndex()); // Store the column index
                    dateHeaders.add(getCellValueAsString(cell).trim()); // Store the original header
                }

                // Store all headers for dynamic processing
                this.lastHeaders.add(getCellValueAsString(cell).trim());
            }

            // Validate that we found at least a Name column and at least 1 date column
            // We can work with just these minimum requirements
            if (nameCol == -1 || dateCols.size() < 1) {
                throw new IOException(
                        "Required columns not found. Excel file must contain at least: Name, and at least 1 date column (e.g., 16-Dec, 17-Dec)");
            }

            // If we don't have Sr No, we'll set it to 0 or derive from row number
            if (srNoCol == -1) {
                System.out.println("Warning: 'Sr No' column not found, will use row numbers");
            }

            // If we have fewer than 2 date columns, we'll work with what we have
            if (dateCols.size() < 2) {
                System.out.println("Warning: Only " + dateCols.size() + " date column(s) found, minimum 2 recommended");
            }

            // Process data rows
            for (int i = headerRow.getRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) {
                    continue; // Skip empty rows
                }

                Employee employee = new Employee();

                // Get Serial Number
                if (srNoCol != -1 && srNoCol < row.getLastCellNum()) {
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
                } else {
                    // If Sr No column not found, use row index as serial number
                    employee.setSrNo(i - headerRow.getRowNum());
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
    /**
     * Gets cell value as string regardless of cell type, using POI's DataFormatter
     * to preserve the format as seen in Excel (crucial for Dates).
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null)
            return "";

        // Use DataFormatter to get the value as shown in Excel
        // This handles dates (e.g., returns "16-Dec" instead of raw numbers)
        // and numbers (e.g., "1" instead of "1.0") correctly.
        return dataFormatter.formatCellValue(cell).trim();
    }

    /**
     * Checks if a row is empty (all cells are blank)
     */
    private boolean isRowEmpty(Row row) {
        if (row == null)
            return true;

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
        if (row == null)
            return false;

        // Count how many cells in this row have meaningful headers
        int headerMatches = 0;
        int totalCells = 0;

        for (Cell cell : row) {
            totalCells++;
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
                headerMatches++;
            }
        }

        // If at least 20% of the cells in this row match known column types, consider
        // it a data header row
        // Lowered threshold to better handle sheets with many columns
        return totalCells > 0 && ((double) headerMatches / totalCells) > 0.2;
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

        // Also check the original header without converting to lowercase for date
        // patterns
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

        // Check for patterns like "16-Dec", "01-Jan", etc., with optional year
        if (origHeader.matches("\\d{1,2}-[A-Za-z]{3}(-[0-9]{2,4})?")) {
            return true;
        }

        // Check for date patterns with spaces like "16 - Dec" or " 16 - Dec "
        String normalizedHeader = lowerHeader.replaceAll("\\s+", "");
        if (normalizedHeader.matches("\\d{1,2}-[a-z]{3}")) {
            return true;
        }

        return false;
    }

    /**
     * Checks if a row contains key headers like Sr No, Name, or date columns
     */
    private boolean containsKeyHeaders(Row row) {
        if (row == null)
            return false;

        boolean hasSrNo = false;
        boolean hasName = false;
        int dateColumns = 0;

        for (Cell cell : row) {
            String header = getCellValueAsString(cell).toLowerCase().trim();

            if (header.contains("sr") && (header.contains("no") || header.contains("num"))) {
                hasSrNo = true;
            } else if (header.contains("name") && !header.contains("post") && !header.contains("pre")) {
                hasName = true;
            } else if (isDateColumn(header)) {
                dateColumns++;
            }
        }

        // Require Sr No and Name, plus at least 1 date column (we'll find another date
        // column in the rest of the method)
        return hasSrNo && hasName && dateColumns >= 1;
    }

    /**
     * Normalizes header text by converting to lowercase and removing special
     * characters
     */
    private String normalizeHeader(String header) {
        if (header == null)
            return "";

        // Convert to lowercase and remove special characters except alphanumeric and
        // spaces
        return header.toLowerCase().replaceAll("[^a-z0-9\\s]", " ").trim();
    }

    /**
     * Creates a formatted Excel sheet with attendance summary in the exact format
     * required
     */
    public void createFormattedAttendanceSummarySheet(AttendanceSummary summary, String day1Date, String day2Date,
            String fileName) throws IOException {
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
        cell.setCellValue("Report Generated On: "
                + java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));

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

    public void insertEmailSummaryIntoExistingSheet(EmailRecord record, String masterFilePath, int startRowNum)
            throws IOException {
        // This method is now deprecated as per requirements.
        // Email summaries should not be stored in the master sheet.
        // The master sheet should only contain tabular employee data.
    }

    /**
     * Adds employee data to the master sheet in tabular format with required
     * columns
     */
    public void addEmployeeDataToMasterSheet(List<Employee> employees, String masterFilePath) throws IOException {
        FileInputStream fis;
        Workbook workbook;

        // Try to load existing workbook, create new one if it doesn't exist
        try {
            fis = new FileInputStream(masterFilePath);
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

        // Check if the header row already exists
        boolean headerExists = false;
        List<String> masterHeaders = new ArrayList<>();

        if (sheet.getLastRowNum() >= 0) {
            Row firstRow = sheet.getRow(0);
            if (firstRow != null && !isRowEmpty(firstRow)) {
                headerExists = true;
                // Read existing master headers
                for (Cell cell : firstRow) {
                    masterHeaders.add(getCellValueAsString(cell).trim());
                }
            }
        }

        int currentRow = 0;

        // Create or find the header row
        if (!headerExists) {
            Row headerRow = sheet.createRow(currentRow++);

            // Use the last headers captured from the source file
            if (this.lastHeaders != null && !this.lastHeaders.isEmpty()) {
                masterHeaders.addAll(this.lastHeaders);
            } else {
                // Fallback if no headers captured (shouldn't happen if readAttendanceData was
                // called)
                // Or we can default to some standard headers
                masterHeaders.add("Sr No");
                masterHeaders.add("Name");
                masterHeaders.add("Trade");
                masterHeaders.add("Date");
            }

            for (int i = 0; i < masterHeaders.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(masterHeaders.get(i));
                cell.setCellStyle(boldStyle);
            }
        } else {
            // If header exists, start from the next available row
            currentRow = sheet.getLastRowNum() + 1;
        }

        // Add employee data to the sheet
        for (Employee emp : employees) {
            Row row = sheet.createRow(currentRow++);

            // For each column in the master sheet, find the value from the employee
            for (int i = 0; i < masterHeaders.size(); i++) {
                String header = masterHeaders.get(i);
                String value = "";

                // Try to get from dynamic fields first (exact match)
                if (emp.getDynamicFields().containsKey(header)) {
                    value = emp.getDynamicFields().get(header);
                } else {
                    // Fallback to standard bean properties if needed, or try case-insensitive match
                    // Since we populated dynamicFields with ALL columns from source, we should find
                    // it if it was in source.
                    // If the master sheet has a column that was NOT in the source, we leave it
                    // blank.
                }

                row.createCell(i).setCellValue(value);
            }
        }

        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(masterFilePath)) {
            workbook.write(fileOut);
        }

        workbook.close();
    }

}