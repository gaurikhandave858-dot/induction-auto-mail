# Automated Attendance Tracking and Email Reporting System

An automated system for processing attendance data from Excel sheets and generating formatted email reports. The system can handle various Excel formats and provides a user-friendly GUI for processing attendance data. It includes a NOT OK classification feature that highlights employees who are absent on one or both days.

## Features

- **Flexible Excel Processing**: Dynamically detects headers and processes various Excel formats with different column layouts
- **Attendance Summary**: Calculates present/absent counts for each day
- **Email Generation**: Creates formatted emails with present/absent statistics and tables of absent students
- **NOT OK Classification**: Highlights employees who are absent on one or both days
- **Tabular Display**: Shows attendance data in structured tables (SR.NO, NAME, DEPARTMENT)
- **Master Sheet Management**: Maintains a central record of employee data in a standardized 19-column format
- **GUI Interface**: User-friendly Swing-based interface with tabbed views for logs, email preview, and table data
- **Filtering Options**: Filter employees by attendance status and absenteeism patterns

## Installation

1. Clone the repository:
```bash
git clone https://github.com/gaurikhandave858-dot/induction-auto-mail.git
```

2. Navigate to the project directory:
```bash
cd induction-auto-mail
```

3. Build the project:
```bash
mvn clean compile package
```

## Usage

1. Run the application:
```bash
mvn exec:java@main
```

2. Select an Excel file containing attendance data
3. Enter recipient email address
4. Click "Process Summary Email" to analyze the data
5. Review the email content in the preview tab
6. Click "Send Email" to send the report

## Excel Format Support

The system supports various Excel formats and automatically detects:
- Employee names (must contain a "Name" column)
- Date columns (formatted as dd-MMM like "16-Dec")
- Additional employee data (P No, Gender, Trade, Mobile No, etc.)

## Master Sheet Structure

Processed data is stored in a master sheet with the following 19-column structure:
- Sr No
- P No
- Name
- Gender
- Batch No
- Grade
- Induction Month
- Induction Start Date
- Induction End Date
- Induction Day 1
- Induction Day 2
- Induction Day 3
- Induction Total HRS
- Pre-Test
- Induction Marks
- Induction Re-Exam
- Induction Total
- Exam Status
- Shop

## Email Format

Emails are generated with:
- Present/absent counts for each day
- Tables showing absent students with SR.NO, NAME, and DEPARTMENT
- NOT OK classification section highlighting employees who are absent on one or both days
- Proper formatting for readability

## Dependencies

- Apache POI (for Excel processing)
- JavaMail API (for email functionality)
- Maven (for build management)

## Build

To build a standalone JAR with all dependencies:
```bash
mvn clean package
```

The resulting JAR file will be located in the `target` directory with all dependencies included.

## License

This project is available for use under the MIT license.