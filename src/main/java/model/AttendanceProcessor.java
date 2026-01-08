package model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AttendanceProcessor {
    
    public AttendanceSummary calculateAttendanceSummary(List<Employee> employees) {
        AttendanceSummary summary = new AttendanceSummary();
        
        List<String> day1AbsentNames = new ArrayList<>();
        List<String> day2AbsentNames = new ArrayList<>();
        
        int day1PresentCount = 0;
        int day1AbsentCount = 0;
        int day2PresentCount = 0;
        int day2AbsentCount = 0;
        
        for (Employee employee : employees) {
            // Process Day 1 attendance
            if (employee.getDay1Attendance() != null && !employee.getDay1Attendance().trim().isEmpty()) {
                if (employee.getDay1Attendance().equalsIgnoreCase("P") || 
                    employee.getDay1Attendance().equalsIgnoreCase("Present")) {
                    day1PresentCount++;
                } else if (employee.getDay1Attendance().equalsIgnoreCase("A") || 
                          employee.getDay1Attendance().equalsIgnoreCase("Absent")) {
                    day1AbsentCount++;
                    day1AbsentNames.add(employee.getName());
                }
            }
            
            // Process Day 2 attendance
            if (employee.getDay2Attendance() != null && !employee.getDay2Attendance().trim().isEmpty()) {
                if (employee.getDay2Attendance().equalsIgnoreCase("P") || 
                    employee.getDay2Attendance().equalsIgnoreCase("Present")) {
                    day2PresentCount++;
                } else if (employee.getDay2Attendance().equalsIgnoreCase("A") || 
                          employee.getDay2Attendance().equalsIgnoreCase("Absent")) {
                    day2AbsentCount++;
                    day2AbsentNames.add(employee.getName());
                }
            }
        }
        
        summary.setDay1PresentCount(day1PresentCount);
        summary.setDay1AbsentCount(day1AbsentCount);
        summary.setDay1AbsentNames(String.join(", ", day1AbsentNames));
        summary.setDay2PresentCount(day2PresentCount);
        summary.setDay2AbsentCount(day2AbsentCount);
        summary.setDay2AbsentNames(String.join(", ", day2AbsentNames));
        
        return summary;
    }
    
    public Map<String, List<Employee>> getAbsentEmployeesByDay(List<Employee> employees) {
        Map<String, List<Employee>> absentEmployees = new HashMap<>();
        absentEmployees.put("day1", new ArrayList<>());
        absentEmployees.put("day2", new ArrayList<>());
        
        for (Employee employee : employees) {
            // Check Day 1
            if (employee.getDay1Attendance() != null && 
                (employee.getDay1Attendance().equalsIgnoreCase("A") || 
                 employee.getDay1Attendance().equalsIgnoreCase("Absent"))) {
                absentEmployees.get("day1").add(employee);
            }
            
            // Check Day 2
            if (employee.getDay2Attendance() != null && 
                (employee.getDay2Attendance().equalsIgnoreCase("A") || 
                 employee.getDay2Attendance().equalsIgnoreCase("Absent"))) {
                absentEmployees.get("day2").add(employee);
            }
        }
        
        return absentEmployees;
    }
    
    public List<Employee> filterEmployeesByStatus(List<Employee> employees, 
                                               AttendanceFilter.AttendanceStatus status, 
                                               String day) {
        return AttendanceFilter.filterByAttendanceStatus(employees, status, day);
    }
    
    public List<Employee> filterEmployeesByAbsenteeism(List<Employee> employees, 
                                                   AttendanceFilter.AbsenteeismType absenteeismType) {
        return AttendanceFilter.filterByAbsenteeismType(employees, absenteeismType);
    }
    
    public int getFilteredCount(List<Employee> employees, 
                                AttendanceFilter.AttendanceStatus status, 
                                String day) {
        return AttendanceFilter.getCountByFilter(employees, status, day);
    }
    
    public int getAbsenteeismCount(List<Employee> employees, 
                                   AttendanceFilter.AbsenteeismType absenteeismType) {
        return AttendanceFilter.getCountByAbsenteeismType(employees, absenteeismType);
    }
    
    public String getAbsenteeismNames(List<Employee> employees, 
                                      AttendanceFilter.AbsenteeismType absenteeismType) {
        return AttendanceFilter.getNamesByAbsenteeismType(employees, absenteeismType);
    }
    
    public String determineOverallStatus(Employee employee) {
        String day1 = employee.getDay1Attendance();
        String day2 = employee.getDay2Attendance();
        
        boolean day1Present = (day1 != null && 
                              (day1.equalsIgnoreCase("P") || day1.equalsIgnoreCase("Present")));
        boolean day2Present = (day2 != null && 
                              (day2.equalsIgnoreCase("P") || day2.equalsIgnoreCase("Present")));
        
        if (day1Present && day2Present) {
            return "Fully Attended";
        } else if (day1Present || day2Present) {
            return "Partially Attended";
        } else {
            return "Not Attended";
        }
    }
}