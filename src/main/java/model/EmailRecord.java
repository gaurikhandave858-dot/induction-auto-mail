package model;

public class EmailRecord {
    private String date;
    private int day1PresentCount;
    private int day1AbsentCount;
    private String day1AbsentNames;
    private int day2PresentCount;
    private int day2AbsentCount;
    private String day2AbsentNames;
    private String emailContent;
    private String day1DateString;
    private String day2DateString;
    
    public EmailRecord() {
    }
    
    public EmailRecord(String date, int day1PresentCount, int day1AbsentCount, String day1AbsentNames,
                       int day2PresentCount, int day2AbsentCount, String day2AbsentNames, String emailContent) {
        this.date = date;
        this.day1PresentCount = day1PresentCount;
        this.day1AbsentCount = day1AbsentCount;
        this.day1AbsentNames = day1AbsentNames;
        this.day2PresentCount = day2PresentCount;
        this.day2AbsentCount = day2AbsentCount;
        this.day2AbsentNames = day2AbsentNames;
        this.emailContent = emailContent;
    }
    
    public EmailRecord(String date, int day1PresentCount, int day1AbsentCount, String day1AbsentNames,
                       int day2PresentCount, int day2AbsentCount, String day2AbsentNames, String emailContent,
                       String day1DateString, String day2DateString) {
        this.date = date;
        this.day1PresentCount = day1PresentCount;
        this.day1AbsentCount = day1AbsentCount;
        this.day1AbsentNames = day1AbsentNames;
        this.day2PresentCount = day2PresentCount;
        this.day2AbsentCount = day2AbsentCount;
        this.day2AbsentNames = day2AbsentNames;
        this.emailContent = emailContent;
        this.day1DateString = day1DateString;
        this.day2DateString = day2DateString;
    }
    
    // Getters and Setters
    public String getDate() {
        return date;
    }
    
    public void setDate(String date) {
        this.date = date;
    }
    
    public int getDay1PresentCount() {
        return day1PresentCount;
    }
    
    public void setDay1PresentCount(int day1PresentCount) {
        this.day1PresentCount = day1PresentCount;
    }
    
    public int getDay1AbsentCount() {
        return day1AbsentCount;
    }
    
    public void setDay1AbsentCount(int day1AbsentCount) {
        this.day1AbsentCount = day1AbsentCount;
    }
    
    public String getDay1AbsentNames() {
        return day1AbsentNames;
    }
    
    public void setDay1AbsentNames(String day1AbsentNames) {
        this.day1AbsentNames = day1AbsentNames;
    }
    
    public int getDay2PresentCount() {
        return day2PresentCount;
    }
    
    public void setDay2PresentCount(int day2PresentCount) {
        this.day2PresentCount = day2PresentCount;
    }
    
    public int getDay2AbsentCount() {
        return day2AbsentCount;
    }
    
    public void setDay2AbsentCount(int day2AbsentCount) {
        this.day2AbsentCount = day2AbsentCount;
    }
    
    public String getDay2AbsentNames() {
        return day2AbsentNames;
    }
    
    public void setDay2AbsentNames(String day2AbsentNames) {
        this.day2AbsentNames = day2AbsentNames;
    }
    
    public String getEmailContent() {
        return emailContent;
    }
    
    public void setEmailContent(String emailContent) {
        this.emailContent = emailContent;
    }
    
    public String getDay1DateString() {
        return day1DateString;
    }
    
    public void setDay1DateString(String day1DateString) {
        this.day1DateString = day1DateString;
    }
    
    public String getDay2DateString() {
        return day2DateString;
    }
    
    public void setDay2DateString(String day2DateString) {
        this.day2DateString = day2DateString;
    }
}