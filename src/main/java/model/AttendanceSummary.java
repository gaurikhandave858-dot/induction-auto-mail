package model;

public class AttendanceSummary {
    private int day1PresentCount;
    private int day1AbsentCount;
    private String day1AbsentNames;
    private int day2PresentCount;
    private int day2AbsentCount;
    private String day2AbsentNames;
    
    // Getters and Setters
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
    
    @Override
    public String toString() {
        return "AttendanceSummary{" +
                "day1PresentCount=" + day1PresentCount +
                ", day1AbsentCount=" + day1AbsentCount +
                ", day1AbsentNames='" + day1AbsentNames + '\'' +
                ", day2PresentCount=" + day2PresentCount +
                ", day2AbsentCount=" + day2AbsentCount +
                ", day2AbsentNames='" + day2AbsentNames + '\'' +
                '}';
    }
}