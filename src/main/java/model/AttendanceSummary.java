package model;

import java.util.ArrayList;
import java.util.List;

public class AttendanceSummary {
    private int day1PresentCount;
    private int day1AbsentCount;
    private String day1AbsentNames;
    private String day1DateString; // Added for enhanced format
    private int day2PresentCount;
    private int day2AbsentCount;
    private String day2AbsentNames;
    private String day2DateString; // Added for enhanced format
    
    // Enhanced fields for multiple date support
    private List<String> allDateHeaders;
    private List<Integer> allDatePresentCounts;
    private List<Integer> allDateAbsentCounts;
    private List<List<String>> allDateAbsentNamesList;
    
    public AttendanceSummary() {
        this.allDateHeaders = new ArrayList<>();
        this.allDatePresentCounts = new ArrayList<>();
        this.allDateAbsentCounts = new ArrayList<>();
        this.allDateAbsentNamesList = new ArrayList<>();
    }
    
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
    
    public String getDay1DateString() {
        return day1DateString;
    }
    
    public void setDay1DateString(String day1DateString) {
        this.day1DateString = day1DateString;
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
    
    public String getDay2DateString() {
        return day2DateString;
    }
    
    public void setDay2DateString(String day2DateString) {
        this.day2DateString = day2DateString;
    }
    
    // Enhanced getters and setters for multiple date support
    public List<String> getAllDateHeaders() {
        return allDateHeaders;
    }
    
    public void setAllDateHeaders(List<String> allDateHeaders) {
        this.allDateHeaders = allDateHeaders;
    }
    
    public List<Integer> getAllDatePresentCounts() {
        return allDatePresentCounts;
    }
    
    public void setAllDatePresentCounts(List<Integer> allDatePresentCounts) {
        this.allDatePresentCounts = allDatePresentCounts;
    }
    
    public List<Integer> getAllDateAbsentCounts() {
        return allDateAbsentCounts;
    }
    
    public void setAllDateAbsentCounts(List<Integer> allDateAbsentCounts) {
        this.allDateAbsentCounts = allDateAbsentCounts;
    }
    
    public List<List<String>> getAllDateAbsentNamesList() {
        return allDateAbsentNamesList;
    }
    
    public void setAllDateAbsentNamesList(List<List<String>> allDateAbsentNamesList) {
        this.allDateAbsentNamesList = allDateAbsentNamesList;
    }
    
    @Override
    public String toString() {
        return "AttendanceSummary{" +
                "day1PresentCount=" + day1PresentCount +
                ", day1AbsentCount=" + day1AbsentCount +
                ", day1AbsentNames='" + day1AbsentNames + '\'' +
                ", day1DateString='" + day1DateString + '\'' +
                ", day2PresentCount=" + day2PresentCount +
                ", day2AbsentCount=" + day2AbsentCount +
                ", day2AbsentNames='" + day2AbsentNames + '\'' +
                ", day2DateString='" + day2DateString + '\'' +
                ", allDateHeaders=" + allDateHeaders +
                ", allDatePresentCounts=" + allDatePresentCounts +
                ", allDateAbsentCounts=" + allDateAbsentCounts +
                ", allDateAbsentNamesList=" + allDateAbsentNamesList +
                '}';
    }
}