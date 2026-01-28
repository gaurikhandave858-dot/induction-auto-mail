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

### Email Configuration - Critical Setup Steps

For Gmail, you MUST follow these steps precisely:

1. **Enable 2-Factor Authentication** on your Gmail account
   - Go to Google Account Settings
   - Navigate to Security → 2-Step Verification
   - Turn on 2-Step Verification

2. **Generate a Gmail App Password**
   - Go to Google Account Settings
   - Navigate to Security → 2-Step Verification → App passwords
   - Select "Mail" as the app
   - Choose your device (can be any, like "Windows Computer")
   - Google will generate a 16-character app password
   - **COPY this password immediately** (it won't be shown again)

3. **Configure the Application**
   - Update `SENDER_EMAIL` with your email address
   - Update `SENDER_APP_PASSWORD` with your 16-character app password (remove spaces)
   - Example: If app password shows as `abcd efgh ijkl mnop`, use `abcdefghijklmnop`

4. Run the application:
```bash
mvn exec:java@main
```

5. Select an Excel file containing attendance data
6. Enter recipient email address
7. Click "Process Summary Email" to analyze the data
8. Review the email content in the preview tab
9. Click "Send Email" to send the report

## Excel Format Support

The system supports various Excel formats and automatically detects:
- Employee names (must contain a "Name" column)
- Date columns (formatted as dd-MMM like "16-Dec")
- Additional employee data (P No, Gender, Trade, Mobile No, etc.)

## Master Sheet Structure

The system separates data persistence and email reporting into two independent tasks:

### Data Storage Workflow
- Attendance data is manually stored into the master Excel sheet using a button embedded in the attendance Excel file.
- Email analysis and reporting are handled exclusively by the Java Swing application and do not modify the master sheet.

### Master Sheet Format
The master sheet maintains the following structure:
- Sr No
- P No
- Name
- Gender
- Batch No
- Grade
- Induction Month
- Induction Start Date
- Induction End Date
- Joining Date *(NEW)*
- Ending Date *(NEW)*
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

This design improves data safety, modularity, and maintainability.

## Email Format

Emails are generated with:
- Present/absent counts for each day
- Tables showing absent students with SR.NO, NAME, and DEPARTMENT
- NOT OK classification section highlighting employees who are absent on one or both days
- Proper formatting for readability

## Troubleshooting

### Email Authentication Issues

If you encounter "Username and Password not accepted" error:
- Verify that 2-factor authentication is enabled on your Gmail account
- Confirm you generated an app password specifically for "Mail"
- Ensure the app password is exactly 16 characters
- Make sure you're using the app password, not your regular Gmail password
- Double-check that the email address in Config.java matches your Gmail account

### Testing Email Configuration

You can test your email configuration separately:
```bash
mvn exec:java@test-email
```

### Validating Credentials

To validate your credentials:
```bash
mvn exec:java@validate-credentials
```

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