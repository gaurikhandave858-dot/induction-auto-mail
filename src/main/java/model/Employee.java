package model;

import java.util.ArrayList;
import java.util.List;

public class Employee {
    private int srNo;
    private String pNo; // P No
    private String name;
    private String gender;
    private String trade;
    private String contactNo;
    private String mobileNo;
    private String preTest;
    private String postTest;
    private String reExam;
    private String deployShop;
    private String day1Attendance; // P for Present, A for Absent
    private String day2Attendance; // P for Present, A for Absent
    private String day1Date;
    private String day2Date;
    private List<String> allDateAttendances;
    private List<String> allDateHeaders;
    
    public Employee() {
        this.allDateAttendances = new ArrayList<>();
        this.allDateHeaders = new ArrayList<>();
    }
    
    public Employee(int srNo, String name, String contactNo, String day1Attendance, String day2Attendance, String day1Date, String day2Date) {
        this.srNo = srNo;
        this.name = name;
        this.contactNo = contactNo;
        this.day1Attendance = day1Attendance;
        this.day2Attendance = day2Attendance;
        this.day1Date = day1Date;
        this.day2Date = day2Date;
        this.allDateAttendances = new ArrayList<>();
        this.allDateHeaders = new ArrayList<>();
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
    
    public String getPNo() {
        return pNo;
    }
    
    public void setPNo(String pNo) {
        this.pNo = pNo;
    }
    
    public String getGender() {
        return gender;
    }
    
    public void setGender(String gender) {
        this.gender = gender;
    }
    
    public String getTrade() {
        return trade;
    }
    
    public void setTrade(String trade) {
        this.trade = trade;
    }
    
    public String getMobileNo() {
        return mobileNo;
    }
    
    public void setMobileNo(String mobileNo) {
        this.mobileNo = mobileNo;
    }
    
    public String getPreTest() {
        return preTest;
    }
    
    public void setPreTest(String preTest) {
        this.preTest = preTest;
    }
    
    public String getPostTest() {
        return postTest;
    }
    
    public void setPostTest(String postTest) {
        this.postTest = postTest;
    }
    
    public String getReExam() {
        return reExam;
    }
    
    public void setReExam(String reExam) {
        this.reExam = reExam;
    }
    
    public String getDeployShop() {
        return deployShop;
    }
    
    public void setDeployShop(String deployShop) {
        this.deployShop = deployShop;
    }
    
    public List<String> getAllDateAttendances() {
        return allDateAttendances;
    }
    
    public void setAllDateAttendances(List<String> allDateAttendances) {
        this.allDateAttendances = allDateAttendances;
    }
    
    public List<String> getAllDateHeaders() {
        return allDateHeaders;
    }
    
    public void setAllDateHeaders(List<String> allDateHeaders) {
        this.allDateHeaders = allDateHeaders;
    }
    
    @Override
    public String toString() {
        return "Employee{" +
                "srNo=" + srNo +
                ", pNo='" + pNo + '\'' +
                ", name='" + name + '\'' +
                ", gender='" + gender + '\'' +
                ", trade='" + trade + '\'' +
                ", contactNo='" + contactNo + '\'' +
                ", mobileNo='" + mobileNo + '\'' +
                ", preTest='" + preTest + '\'' +
                ", postTest='" + postTest + '\'' +
                ", reExam='" + reExam + '\'' +
                ", deployShop='" + deployShop + '\'' +
                ", day1Attendance='" + day1Attendance + '\'' +
                ", day2Attendance='" + day2Attendance + '\'' +
                ", day1Date='" + day1Date + '\'' +
                ", day2Date='" + day2Date + '\'' +
                ", allDateAttendances=" + allDateAttendances +
                ", allDateHeaders=" + allDateHeaders +
                '}';
    }
}