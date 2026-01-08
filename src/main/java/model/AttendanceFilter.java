package model;

import java.util.List;
import java.util.stream.Collectors;

public class AttendanceFilter {
    
    public enum AttendanceStatus {
        PRESENT, ABSENT, ALL
    }
    
    public enum AbsenteeismType {
        ABSENT_ON_DAY1, ABSENT_ON_DAY2, ABSENT_ON_BOTH_DAYS, ABSENT_ON_ANY_DAY, ALL
    }
    
    /**
     * Filter employees by attendance status for a specific day
     */
    public static List<Employee> filterByAttendanceStatus(List<Employee> employees, 
                                                         AttendanceStatus status, 
                                                         String day) {
        if (employees == null) return null;
        
        return employees.stream()
            .filter(emp -> {
                String attendance = day.equalsIgnoreCase("day1") ? 
                    emp.getDay1Attendance() : emp.getDay2Attendance();
                
                if (attendance == null) return false;
                
                boolean isPresent = attendance.equalsIgnoreCase("P") || 
                                  attendance.equalsIgnoreCase("Present");
                boolean isAbsent = attendance.equalsIgnoreCase("A") || 
                                 attendance.equalsIgnoreCase("Absent");
                
                switch (status) {
                    case PRESENT:
                        return isPresent;
                    case ABSENT:
                        return isAbsent;
                    case ALL:
                    default:
                        return true;
                }
            })
            .collect(Collectors.toList());
    }
    
    /**
     * Filter employees by absenteeism type
     */
    public static List<Employee> filterByAbsenteeismType(List<Employee> employees, 
                                                        AbsenteeismType absenteeismType) {
        if (employees == null) return null;
        
        return employees.stream()
            .filter(emp -> {
                boolean day1Absent = isAbsent(emp.getDay1Attendance());
                boolean day2Absent = isAbsent(emp.getDay2Attendance());
                
                switch (absenteeismType) {
                    case ABSENT_ON_DAY1:
                        return day1Absent && !day2Absent;
                    case ABSENT_ON_DAY2:
                        return day2Absent && !day1Absent;
                    case ABSENT_ON_BOTH_DAYS:
                        return day1Absent && day2Absent;
                    case ABSENT_ON_ANY_DAY:
                        return day1Absent || day2Absent;
                    case ALL:
                    default:
                        return true;
                }
            })
            .collect(Collectors.toList());
    }
    
    /**
     * Helper method to check if attendance status indicates absence
     */
    private static boolean isAbsent(String attendance) {
        if (attendance == null) return false;
        return attendance.equalsIgnoreCase("A") || 
               attendance.equalsIgnoreCase("Absent");
    }
    
    /**
     * Get count of employees matching the filter criteria
     */
    public static int getCountByFilter(List<Employee> employees, 
                                      AttendanceStatus status, 
                                      String day) {
        List<Employee> filtered = filterByAttendanceStatus(employees, status, day);
        return filtered != null ? filtered.size() : 0;
    }
    
    /**
     * Get count of employees by absenteeism type
     */
    public static int getCountByAbsenteeismType(List<Employee> employees, 
                                               AbsenteeismType absenteeismType) {
        List<Employee> filtered = filterByAbsenteeismType(employees, absenteeismType);
        return filtered != null ? filtered.size() : 0;
    }
    
    /**
     * Get names of employees matching the absenteeism filter
     */
    public static String getNamesByAbsenteeismType(List<Employee> employees, 
                                                  AbsenteeismType absenteeismType) {
        List<Employee> filtered = filterByAbsenteeismType(employees, absenteeismType);
        if (filtered == null || filtered.isEmpty()) {
            return "";
        }
        
        return filtered.stream()
            .map(Employee::getName)
            .collect(Collectors.joining(", "));
    }
}