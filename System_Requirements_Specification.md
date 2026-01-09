# System Requirements Specification (SRS)
## Enhanced Attendance Excel Analysis System

---

## Table of Contents
1. [Introduction](#introduction)
2. [Supported Excel Sheet Structure](#supported-excel-sheet-structure)
3. [Functional Requirements](#functional-requirements)
4. [Non-Functional Requirements](#non-functional-requirements)
5. [Assumptions](#assumptions)
6. [Future Enhancements](#future-enhancements)

---

## 1. Introduction

The Enhanced Attendance Excel Analysis System is designed to process attendance data from Excel sheets used in employee induction programs. The system analyzes employee attendance across multiple training days, generates summary reports, and maintains historical records in a master Excel file.

### 1.1 Purpose
This document outlines the functional and non-functional requirements for the Enhanced Attendance Excel Analysis System, focusing on dynamic Excel processing, attendance calculation, email generation, and master data maintenance.

### 1.2 Scope
The system processes Excel files containing employee attendance data from induction programs, calculates attendance statistics, generates email summaries, and maintains historical records.

---

## 2. Supported Excel Sheet Structure

The system shall support analysis of Excel sheets formatted similar to the uploaded induction attendance sheet, which includes:

### 2.1 Header Section (Informational)
Training modules listed at the top (e.g.):
- Safety Training
- POSH & Health
- Workplace Discipline & TCoC
- Quality Orientation
- Bus Pass
- Time Office
- Self Directed Teams
- 5S
- Induction Training Post Test / Feedback

These rows are non-data rows and shall be ignored during attendance processing.

### 2.2 Employee Attendance Data Section
The system shall dynamically detect the attendance table starting row and process the following columns:

| Column Name | Description |
|-------------|-------------|
| Sr No | Serial Number |
| P No | Employee Personnel Number |
| Name | Employee Name |
| Gender | Male / Female |
| Trade | Employee Trade |
| Date Columns (e.g. 16-Dec, 17-Dec) | Attendance marked as P or A |
| Mobile No | Contact Number |
| Pre-Test | Assessment Score (optional) |
| Post-Test | Assessment Score (optional) |
| Re-Exam | Assessment Status (optional) |
| Deploy Shop | Deployment Information |

---

## 3. Functional Requirements

### 3.1 Excel Format Detection Module (New)
The system shall:
- Automatically identify the attendance table section in the Excel file.
- Ignore header/training description rows.
- Detect date-based attendance columns dynamically (e.g., 16-Dec, 17-Dec).
- No hardcoded column index shall be used for attendance dates.

### 3.2 Attendance Analysis Module (Updated)
For each detected date column:
- Calculate Total Present count
- Calculate Total Absent count
- Extract names of employees marked A
- Support any number of induction days (minimum 2)

Attendance values supported:
- P → Present
- A → Absent

### 3.3 Summary Email Generation (Updated)
The system shall generate an email summary in the following format:

#### Example:
```
EMAIL SUMMARY CONTENT

Report Generated On: 09/01/2026

Day 1:

Day 1 (16-Dec-2025) Summary:
Present: 6
Absent: 2

Day 1 (16-Dec-2025) Absent Employee List:

- Bhausahab Gangadhar Salve
- Niraj Netaji Bhatlawande

Day 2 (17-Dec-2025) Summary:
Present: 6
Absent: 2

Day 2 (17-Dec-2025) Absent Employee List:

- Bhausahab Gangadhar Salve
- Niraj Netaji Bhatlawande
```

### 3.4 Master Excel Update Module (Updated)
The built-in Master Excel file shall:
- Store attendance summaries derived from this Excel format.
- Maintain historical summaries in email-text format.
- Append new records without overwriting previous entries.

Master Excel shall support:
- Multiple dates
- Multiple batches
- Different induction formats

---

## 4. Non-Functional Requirements

### 4.1 Flexibility
System shall support Excel sheets with:
- Variable header rows
- Dynamic date columns
- Optional assessment columns

### 4.2 Robustness
System shall handle missing or empty optional columns without failure.

### 4.3 Performance
- Process Excel files within reasonable timeframes (under 30 seconds for files with up to 1000 employees)
- Minimal memory usage during processing
- Efficient handling of large Excel files

### 4.4 Compatibility
- Support both .xls and .xlsx file formats
- Compatible with Java 11 and above
- Cross-platform compatibility

---

## 5. Assumptions

- Attendance columns contain only P or A.
- Date columns are uniquely identifiable by date formatting (e.g. dd-MMM).
- Excel files follow the general structure described above.
- Network connectivity is available for email sending functionality.

---

## 6. Future Enhancements

- Auto-detection of training completion status.
- Attendance-to-deployment correlation.
- Graphical attendance trend reports per trade or shop.
- Integration with HR systems for automated data updates.
- Dashboard visualization of attendance trends.
- Export capabilities to multiple formats (PDF, CSV).

---

## Appendix A: Technical Specifications

### A.1 Dependencies
- Apache POI (5.2.3) - For Excel processing
- JavaMail API (1.6.2) - For email functionality
- JCalendar (1.4) - For date handling

### A.2 Architecture Components
- **Model Layer**: Data processing and business logic
- **Controller Layer**: Business logic coordination
- **View Layer**: User interface components
- **Utility Classes**: Excel processing and email sending
