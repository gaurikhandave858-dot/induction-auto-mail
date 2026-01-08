package controller;

import java.io.IOException;
import java.util.List;

import model.AttendanceFilter;
import model.AttendanceProcessor;
import model.AttendanceSummary;
import model.EmailRecord;
import model.EmailSender;
import model.Employee;
import model.ExcelProcessor;

public class AttendanceController {
    private List<Employee> employees;
    private AttendanceProcessor attendanceProcessor;
    private EmailSender emailSender;
    private ExcelProcessor excelProcessor;
    private AttendanceSummary attendanceSummary;
    private String emailContent;
    
    // Email configuration - these would typically come from a config file or user input
    private String smtpHost = "smtp.gmail.com"; // Default, should be configurable
    private String smtpPort = "587";           // Default, should be configurable
    private static final String SENDER_EMAIL = "gaurikhandave858@gmail.com";
    private static final String SENDER_APP_PASSWORD = "xyflmraxlagbwiej";
    private String senderEmail = SENDER_EMAIL;
    private String senderPassword = SENDER_APP_PASSWORD;
    
    public AttendanceController() {
        this.attendanceProcessor = new AttendanceProcessor();
        this.excelProcessor = new ExcelProcessor();
    }
    
    public List<Employee> processAttendanceFile(String filePath) throws IOException {
        this.employees = excelProcessor.readAttendanceData(filePath);
        this.attendanceSummary = attendanceProcessor.calculateAttendanceSummary(employees);
        
        // Generate email content with dates from the Excel file
        if (!employees.isEmpty()) {
            String day1Date = employees.get(0).getDay1Date();
            String day2Date = employees.get(0).getDay2Date();
            
            this.emailSender = new EmailSender(smtpHost, smtpPort, senderEmail, senderPassword);
            this.emailContent = emailSender.generateAttendanceSummaryEmailWithDates(attendanceSummary, day1Date, day2Date);
        }
        
        return this.employees;
    }
    
    public String generateEmailSummary(List<Employee> employees) {
        this.attendanceSummary = attendanceProcessor.calculateAttendanceSummary(employees);
        
        if (!employees.isEmpty()) {
            String day1Date = employees.get(0).getDay1Date();
            String day2Date = employees.get(0).getDay2Date();
            
            this.emailSender = new EmailSender(smtpHost, smtpPort, senderEmail, senderPassword);
            this.emailContent = emailSender.generateAttendanceSummaryEmailWithDates(attendanceSummary, day1Date, day2Date);
        }
        
        return this.emailContent;
    }
    
    public boolean sendEmail(String recipientEmail) {
        if (emailContent == null) {
            System.err.println("Email not prepared. Call processAttendanceFile first.");
            return false;
        }
        
        // Using configured email credentials
        EmailSender tempSender = new EmailSender(smtpHost, smtpPort, senderEmail, senderPassword);
        return tempSender.sendEmail(recipientEmail, "Daily Attendance Summary Report", emailContent);
    }
    
    public boolean sendEmailToSelf() {
        if (emailContent == null || senderEmail.isEmpty()) {
            System.err.println("Self email not configured or email not prepared.");
            return false;
        }
        
        EmailSender tempSender = new EmailSender(smtpHost, smtpPort, senderEmail, senderPassword);
        return tempSender.sendEmail(senderEmail, "Daily Attendance Summary Report (Self Copy)", emailContent);
    }
    
    public void updateMasterSheet() throws IOException {
        if (attendanceSummary != null && emailContent != null) {
            // Create an email record for the master sheet
            EmailRecord record = new EmailRecord();
            record.setDate(java.time.LocalDate.now().toString());
            record.setDay1PresentCount(attendanceSummary.getDay1PresentCount());
            record.setDay1AbsentCount(attendanceSummary.getDay1AbsentCount());
            record.setDay1AbsentNames(attendanceSummary.getDay1AbsentNames());
            record.setDay2PresentCount(attendanceSummary.getDay2PresentCount());
            record.setDay2AbsentCount(attendanceSummary.getDay2AbsentCount());
            record.setDay2AbsentNames(attendanceSummary.getDay2AbsentNames());
            record.setEmailContent(emailContent);
            
            // Set date strings if available from the first employee
            if (!employees.isEmpty()) {
                record.setDay1DateString(employees.get(0).getDay1Date());
                record.setDay2DateString(employees.get(0).getDay2Date());
            }
            
            // Insert the email summary in the exact format starting from row 10 in Column A
            excelProcessor.insertEmailSummaryIntoExistingSheet(record, "master_attendance_summary.xlsx", 10);
        }
    }
    
    // Getters for GUI to access counts
    public int getDay1PresentCount() {
        return attendanceSummary != null ? attendanceSummary.getDay1PresentCount() : 0;
    }
    
    public int getDay1AbsentCount() {
        return attendanceSummary != null ? attendanceSummary.getDay1AbsentCount() : 0;
    }
    
    public int getDay2PresentCount() {
        return attendanceSummary != null ? attendanceSummary.getDay2PresentCount() : 0;
    }
    
    public int getDay2AbsentCount() {
        return attendanceSummary != null ? attendanceSummary.getDay2AbsentCount() : 0;
    }
    
    public String getSelfEmail() {
        return senderEmail;
    }
    
    // Filter methods
    public List<Employee> filterEmployeesByStatus(AttendanceFilter.AttendanceStatus status, String day) {
        if (employees == null) return null;
        return attendanceProcessor.filterEmployeesByStatus(employees, status, day);
    }
    
    public List<Employee> filterEmployeesByAbsenteeism(AttendanceFilter.AbsenteeismType absenteeismType) {
        if (employees == null) return null;
        return attendanceProcessor.filterEmployeesByAbsenteeism(employees, absenteeismType);
    }
    
    public int getFilteredCount(AttendanceFilter.AttendanceStatus status, String day) {
        if (employees == null) return 0;
        return attendanceProcessor.getFilteredCount(employees, status, day);
    }
    
    public int getAbsenteeismCount(AttendanceFilter.AbsenteeismType absenteeismType) {
        if (employees == null) return 0;
        return attendanceProcessor.getAbsenteeismCount(employees, absenteeismType);
    }
    
    public String getAbsenteeismNames(AttendanceFilter.AbsenteeismType absenteeismType) {
        if (employees == null) return "";
        return attendanceProcessor.getAbsenteeismNames(employees, absenteeismType);
    }
    
    // Methods to set email configuration
    public void setEmailConfig(String smtpHost, String smtpPort, String senderEmail, String senderPassword) {
        this.smtpHost = smtpHost;
        this.smtpPort = smtpPort;
        this.senderEmail = senderEmail;
        this.senderPassword = senderPassword;
    }
}