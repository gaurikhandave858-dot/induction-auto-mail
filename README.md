# Automated Attendance Tracking and Email Reporting System

## Overview
This Java-based desktop application automates the process of tracking employee attendance during a two-day induction program and generates automated summary emails with attendance counts. The system is designed to reduce manual effort, minimize errors, and provide professional summary reports.

## Features

### Core Functionality
- **Excel File Processing**: Reads attendance data from Excel files (.xls and .xlsx) with flexible column detection
- **Attendance Calculation**: Automatically calculates present/absent counts for each day
- **Email Generation**: Creates formatted summary emails with dates in brackets (e.g., Day 1 (05/01/2026))
- **Email Sending**: Sends summary emails to designated recipients using JavaMail API
- **Master Sheet Management**: Updates a built-in Excel sheet with all summary data, preserving historical records
- **Self Email**: Sends a copy of the summary email to the sender's email address

### Advanced Features
- **Attendance Status Filter**: Filter employees by attendance status (Present, Absent) for Day 1 or Day 2
- **Absenteeism Filter**: Filter by absenteeism patterns (Absent on Day 1 Only, Day 2 Only, Both Days, or Any Day)
- **Enhanced UI**: Large, visible interface elements for better user experience
- **Confirmation Dialog**: Separate yes/no confirmation before sending emails

## System Requirements
- Java SE 11 or higher
- Apache POI library (for Excel processing)
- JavaMail API (for email functionality)
- At least 4GB RAM recommended

## Installation

1. Ensure you have Java 11+ installed on your system
2. Clone or download this project to your local machine
3. Navigate to the project directory
4. Run the following command to compile the project:
   ```
   mvn compile
   ```
5. Run the application:
   ```
   java -cp "target/classes;target/dependency/*" Main
   ```

Alternative way to run the application:
   ```
   mvn exec:java
   ```

## Configuration

The application is pre-configured with the following email settings:
- **Sender Email**: gaurikhandave858@gmail.com
- **App Password**: xyflmraxlagbwiej
- **SMTP Server**: smtp.gmail.com
- **Port**: 587

*Note: For security reasons, these credentials are embedded in the code. In a production environment, these should be stored securely.*

## Usage

### Basic Workflow
1. Click "Browse Attendance File" to select an Excel file containing attendance data
   - File should contain columns: Sr. No, Name, Contact No, Day 1 (with date), Day 2 (with date)
2. Enter the recipient email address in the "Recipient Email" field
3. Click "Process Summary Email" to analyze the data and generate the summary
4. Review the generated summary in the results panel
5. Use the filter options to analyze specific attendance patterns if needed
6. Click "Send Email" to send the summary (with confirmation dialog)
7. The system will automatically update the master sheet with the summary data

### Using Filters
1. After processing an attendance file, use the filter controls at the bottom
2. **Attendance Status Filter**:
   - Select status (ALL/PRESENT/ABSENT)
   - Select day (Day 1/Day 2)
   - Click "Apply Status Filter"
3. **Absenteeism Filter**:
   - Select absenteeism type from the dropdown
   - Click "Apply Absenteeism Filter"
4. Results will appear in the results text area

## File Structure
- `Main.java`: Application entry point
- `AttendanceEmailFrame.java`: Main GUI interface
- `AttendanceController.java`: Business logic controller
- `Employee.java`: Employee data model
- `AttendanceProcessor.java`: Attendance calculation logic
- `ExcelProcessor.java`: Excel file processing
- `EmailSender.java`: Email generation and sending
- `AttendanceFilter.java`: Filtering functionality
- `master_attendance_summary.xlsx`: Built-in master sheet for historical data

## Email Format
The system generates emails in the following format:

```
EMAIL SUMMARY CONTENT

Report Generated On: dd/MM/yyyy

Day 1:
Day 1 (date) Summary:
Present: [number]
Absent: [number]

Day 1 (date) Absent Employee List:

- Employee Name 1
- Employee Name 2

Day 2 (date) Summary:
Present: [number]
Absent: [number]

Day 2 (date) Absent Employee List:

- Employee Name 1
```

## Master Sheet
The system maintains a single master Excel sheet (`master_attendance_summary.xlsx`) that:
- Preserves all historical attendance summary data
- Automatically updates when new summaries are processed
- Contains formatted email content starting from a specific row (Row 10) in Column A
- Maintains all previous data while appending new entries
- Uses bold formatting for headers and proper spacing for readability

## Security Considerations
- Email credentials are currently stored in the source code for demonstration purposes
- In a production environment, credentials should be stored securely (e.g., encrypted config file, environment variables)
- Ensure your email account has "App Passwords" enabled for programmatic access

## Troubleshooting
- If you encounter issues with email sending, verify that the app password is correct
- Ensure the Excel file follows the expected format with proper column headers
- Check that your system has internet access for sending emails

## License
This project is created for educational purposes as part of an internship project.