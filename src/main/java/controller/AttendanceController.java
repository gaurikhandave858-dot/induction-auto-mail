package controller;

import java.io.IOException;
import java.util.List;

import model.AttendanceFilter;
import model.AttendanceProcessor;
import model.AttendanceSummary;
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
    
    private String smtpHost = model.Config.getSmtpHost();
    private String smtpPort = String.valueOf(model.Config.getSmtpPort());
    private String senderEmail = model.Config.getSenderEmail();
    private String senderPassword = model.Config.getSenderAppPassword();
    
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
            this.emailContent = emailSender.generateAttendanceSummaryEmailWithDates(attendanceSummary, day1Date, day2Date, employees);
        }
        
        return this.employees;
    }
    
    public String generateEmailSummary(List<Employee> employees) {
        this.attendanceSummary = attendanceProcessor.calculateAttendanceSummary(employees);
        
        if (!employees.isEmpty()) {
            String day1Date = employees.get(0).getDay1Date();
            String day2Date = employees.get(0).getDay2Date();
            
            this.emailSender = new EmailSender(smtpHost, smtpPort, senderEmail, senderPassword);
            this.emailContent = emailSender.generateAttendanceSummaryEmailWithDates(attendanceSummary, day1Date, day2Date, employees);
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
    
    
    /**
     * This method has been deprecated as per new system design.
     * Master sheet storage is now handled separately via Excel button.
     * The Java application now focuses only on email analysis and reporting.
     */
    @Deprecated
    public void updateMasterSheet() throws IOException {
        // This functionality has been moved to Excel VBA macro
        // The Java application now focuses only on email processing
        System.out.println("Master sheet storage is now handled via Excel button.");
        System.out.println("This method is deprecated in favor of Excel-based storage.");
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
        model.Config.setSmtpHost(smtpHost);
        model.Config.setSmtpPort(Integer.parseInt(smtpPort));
        model.Config.setSenderEmail(senderEmail);
        model.Config.setSenderAppPassword(senderPassword);
        
        // Update instance variables to reflect the changes
        this.smtpHost = smtpHost;
        this.smtpPort = smtpPort;
        this.senderEmail = senderEmail;
        this.senderPassword = senderPassword;
    }
    
    // Getter methods for email configuration
    public String getSmtpHost() {
        return smtpHost;
    }
    
    public String getSmtpPort() {
        return smtpPort;
    }
    
    public String getSenderEmail() {
        return senderEmail;
    }
}