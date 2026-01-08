package model;

public class Employee {
    private int srNo;
    private String name;
    private String contactNo;
    private String day1Attendance; // P for Present, A for Absent
    private String day2Attendance; // P for Present, A for Absent
    private String day1Date;
    private String day2Date;
    
    public Employee() {
    }
    
    public Employee(int srNo, String name, String contactNo, String day1Attendance, String day2Attendance, String day1Date, String day2Date) {
        this.srNo = srNo;
        this.name = name;
        this.contactNo = contactNo;
        this.day1Attendance = day1Attendance;
        this.day2Attendance = day2Attendance;
        this.day1Date = day1Date;
        this.day2Date = day2Date;
    }
    
    // Getters and Setters
    public int getSrNo() {
        return srNo;
    }
    
    public void setSrNo(int srNo) {
        this.srNo = srNo;
    }
    
    public String getName() {
        return name;
    }
    
    public void setName(String name) {
        this.name = name;
    }
    
    public String getContactNo() {
        return contactNo;
    }
    
    public void setContactNo(String contactNo) {
        this.contactNo = contactNo;
    }
    
    public String getDay1Attendance() {
        return day1Attendance;
    }
    
    public void setDay1Attendance(String day1Attendance) {
        this.day1Attendance = day1Attendance;
    }
    
    public String getDay2Attendance() {
        return day2Attendance;
    }
    
    public void setDay2Attendance(String day2Attendance) {
        this.day2Attendance = day2Attendance;
    }
    
    public String getDay1Date() {
        return day1Date;
    }
    
    public void setDay1Date(String day1Date) {
        this.day1Date = day1Date;
    }
    
    public String getDay2Date() {
        return day2Date;
    }
    
    public void setDay2Date(String day2Date) {
        this.day2Date = day2Date;
    }
    
    @Override
    public String toString() {
        return "Employee{" +
                "srNo=" + srNo +
                ", name='" + name + '\'' +
                ", contactNo='" + contactNo + '\'' +
                ", day1Attendance='" + day1Attendance + '\'' +
                ", day2Attendance='" + day2Attendance + '\'' +
                ", day1Date='" + day1Date + '\'' +
                ", day2Date='" + day2Date + '\'' +
                '}';
    }
}