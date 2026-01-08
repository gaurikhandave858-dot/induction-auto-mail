# Automated Attendance Tracking and Email Reporting System - File Descriptions

## Project Structure
```
src/
├── main/
│   └── java/
│       ├── Main.java
│       ├── TestSystem.java
│       ├── controller/
│       │   └── AttendanceController.java
│       ├── model/
│       │   ├── Employee.java
│       │   ├── EmailRecord.java
│       │   ├── ExcelProcessor.java
│       │   ├── AttendanceProcessor.java
│       │   └── EmailSender.java
│       └── view/
│           ├── AttendanceEmailFrame.java
│           └── ConfirmationDialog.java
└── pom.xml
```

## File Descriptions

### Main Application Entry Point
- **[Main.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/Main.java)**: Main class that launches the GUI application. Sets system look and feel and initializes the main frame.

### Model Layer (Data and Business Logic)
- **[Employee.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/model/Employee.java)**: Represents an employee with properties for serial number, name, contact number, and attendance for Day 1 and Day 2, including date information.

- **[EmailRecord.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/model/EmailRecord.java)**: Represents a record for storing email summaries in the master Excel sheet with counts and names.

- **[ExcelProcessor.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/model/ExcelProcessor.java)**: Handles reading attendance data from Excel files using Apache POI and updating the master Excel sheet with historical data.

- **[AttendanceProcessor.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/model/AttendanceProcessor.java)**: Contains logic to calculate attendance summaries, count present/absent employees, and determine overall attendance status.

- **[EmailSender.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/model/EmailSender.java)**: Manages email generation and sending functionality using JavaMail API, including formatting emails with date brackets as required.

### Controller Layer (Application Logic)
- **[AttendanceController.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/controller/AttendanceController.java)**: Orchestrates the entire process by coordinating between the model components and handling business logic flows.

### View Layer (GUI Interface)
- **[AttendanceEmailFrame.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/view/AttendanceEmailFrame.java)**: Main GUI window with components for browsing attendance files, entering recipient email, processing data, and sending emails.

- **[ConfirmationDialog.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/view/ConfirmationDialog.java)**: Separate dialog for yes/no confirmation when sending emails, as required in the specifications.

### Configuration and Utilities
- **[pom.xml](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/pom.xml)**: Maven configuration file defining project dependencies (Apache POI, JavaMail API, JCalendar) and build settings.

- **[TestSystem.java](file:///C:/Users/gauriGAURI/OneDrive/induction%20auto%20mail(2)/src/main/java/TestSystem.java)**: Test class to verify system components and demonstrate the complete workflow.

## Key Features Implemented

1. **Excel File Processing**: Reads attendance data from Excel files with flexible column detection
2. **Attendance Calculation**: Counts present/absent employees for each day
3. **Email Generation**: Creates formatted emails with dates in brackets (e.g., Day 1 (05/01/2026))
4. **GUI Interface**: Complete user interface with browse, process, and send email functionality
5. **Master Sheet Management**: Updates Excel with historical data preservation
6. **Confirmation Dialog**: Separate yes/no confirmation page as specified
7. **Date Handling**: Properly extracts and displays dates from Excel headers
8. **Absent Employee Listing**: Shows names of absent employees in the summary

## How to Run
Compile and run the application using Maven:
```
mvn clean compile exec:java -Dexec.mainClass="Main"
```