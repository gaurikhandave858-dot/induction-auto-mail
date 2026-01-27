package model;

import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class EmailSender {

    private String smtpHost;
    private String smtpPort;
    private String email;
    private String password;

    public EmailSender(String smtpHost, String smtpPort, String email, String password) {
        this.smtpHost = smtpHost;
        this.smtpPort = smtpPort;
        this.email = email;
        this.password = password;
    }

    public boolean sendEmail(String recipientEmail, String subject, String body) {
        try {
            // Setup mail server properties
            Properties props = new Properties();
            props.put("mail.smtp.host", smtpHost);
            props.put("mail.smtp.port", smtpPort);
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.starttls.enable", "true"); // Enable TLS
            
            // Additional properties for better debugging
            props.put("mail.smtp.connectiontimeout", "10000"); // 10 seconds
            props.put("mail.smtp.timeout", "10000"); // 10 seconds
            props.put("mail.smtp.writetimeout", "10000"); // 10 seconds

            // Create session with authentication
            Session session = Session.getInstance(props, new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(email, password);
                }
            });

            // Validate email addresses before sending
            if (email == null || email.trim().isEmpty()) {
                System.err.println("Sender email is not configured");
                return false;
            }
            
            if (password == null || password.trim().isEmpty()) {
                System.err.println("Sender password is not configured");
                return false;
            }
            
            if (recipientEmail == null || recipientEmail.trim().isEmpty()) {
                System.err.println("Recipient email is not provided");
                return false;
            }

            // Create message
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(email));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject(subject);

            // Set content as HTML
            message.setContent(body, "text/html; charset=utf-8");

            // Send message
            Transport.send(message);

            System.out.println("Email sent successfully to: " + recipientEmail);
            return true;
        } catch (MessagingException e) {
            System.err.println("MessagingException occurred while sending email:");
            System.err.println("  Exception Type: " + e.getClass().getSimpleName());
            System.err.println("  Message: " + e.getMessage());
            
            // Check for specific error types
            Throwable cause = e.getCause();
            if (cause != null) {
                System.err.println("  Cause: " + cause.getMessage());
            }
            
            // Print stack trace for detailed debugging
            e.printStackTrace();
            return false;
        } catch (Exception e) {
            System.err.println("General exception occurred while sending email:");
            System.err.println("  Exception Type: " + e.getClass().getSimpleName());
            System.err.println("  Message: " + e.getMessage());
            e.printStackTrace();
            return false;
        }
    }

    public String generateAttendanceSummaryEmail(AttendanceSummary summary) {
        StringBuilder emailContent = new StringBuilder();

        emailContent.append("<html><body>");
        emailContent.append("<h3>Daily Attendance Summary Report</h3>");
        emailContent.append("<p>Dear Team,</p>");
        emailContent.append("<p>Please find the attendance summary for the induction program:</p>");

        // Day 1 Summary
        emailContent.append("<h4>Day 1 Summary:</h4>");
        emailContent.append("<ul>");
        emailContent.append("<li>Present: ").append(summary.getDay1PresentCount()).append("</li>");
        emailContent.append("<li>Absent: ").append(summary.getDay1AbsentCount()).append("</li>");
        if (!summary.getDay1AbsentNames().isEmpty()) {
            emailContent.append("<li>Absent Names: ").append(summary.getDay1AbsentNames()).append("</li>");
        }
        emailContent.append("</ul>");

        // Day 2 Summary
        emailContent.append("<h4>Day 2 Summary:</h4>");
        emailContent.append("<ul>");
        emailContent.append("<li>Present: ").append(summary.getDay2PresentCount()).append("</li>");
        emailContent.append("<li>Absent: ").append(summary.getDay2AbsentCount()).append("</li>");
        if (!summary.getDay2AbsentNames().isEmpty()) {
            emailContent.append("<li>Absent Names: ").append(summary.getDay2AbsentNames()).append("</li>");
        }
        emailContent.append("</ul>");

        emailContent.append("<p>Best regards,<br>Automated Attendance System</p>");
        emailContent.append("</body></html>");

        return emailContent.toString();
    }

    // Method to generate email with the exact format as specified in SRS
    public String generateAttendanceSummaryEmailWithDates(AttendanceSummary summary, String day1Date, String day2Date,
            List<Employee> employees) {
        StringBuilder emailContent = new StringBuilder();

        emailContent.append("<html><body style='font-family: Arial, sans-serif;'>");
        emailContent.append("<h2>EMAIL SUMMARY CONTENT</h2>");

        String currentDate = java.time.LocalDate.now()
                .format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        emailContent.append("<p><strong>Report Generated On:</strong> ").append(currentDate).append("</p><br>");

        // CSS for table
        String tableStyle = "border-collapse: collapse; width: 100%; border: 1px solid #ddd;";
        String thStyle = "background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: left;";
        String tdStyle = "border: 1px solid #ddd; padding: 8px;";

        // Check if we have multiple date data to process
        if (!summary.getAllDateHeaders().isEmpty()) {
            // Process all date columns in the enhanced format
            List<String> allDateHeaders = summary.getAllDateHeaders();
            List<Integer> presentCounts = summary.getAllDatePresentCounts();
            List<Integer> absentCounts = summary.getAllDateAbsentCounts();
            List<List<String>> absentNamesList = summary.getAllDateAbsentNamesList();

            for (int i = 0; i < allDateHeaders.size(); i++) {
                String dateHeader = allDateHeaders.get(i);
                int presentCount = (i < presentCounts.size()) ? presentCounts.get(i) : 0;
                int absentCount = (i < absentCounts.size()) ? absentCounts.get(i) : 0;
                List<String> absentNames = (i < absentNamesList.size()) ? absentNamesList.get(i) : new ArrayList<>();

                // Format the date properly
                String formattedDate = formatDate(dateHeader);

                // Add day header
                emailContent.append("<h3>Day ").append(i + 1).append(":</h3>");

                // Add summary for this day
                emailContent.append("<p><strong>Day ").append(i + 1).append(" (").append(formattedDate)
                        .append(") Summary:</strong><br>");
                emailContent.append("Present: ").append(presentCount).append("<br>");
                emailContent.append("Absent: ").append(absentCount).append("</p>");

                // Add absent employee list header for this day
                emailContent.append("<h4>Day ").append(i + 1).append(" (").append(formattedDate)
                        .append(") Absent Employee List:</h4>");

                if (!absentNames.isEmpty()) {
                    emailContent.append("<table style='").append(tableStyle).append("'>");
                    emailContent.append("<tr>");
                    emailContent.append("<th style='").append(thStyle).append("'>Sr. No</th>");
                    emailContent.append("<th style='").append(thStyle).append("'>Name</th>");
                    emailContent.append("<th style='").append(thStyle).append("'>Trade/Department</th>");
                    emailContent.append("</tr>");

                    int srNo = 1;
                    for (String name : absentNames) {
                        if (name != null && !name.trim().isEmpty()) {
                            // Find employee trade
                            String trade = "N/A";
                            for (Employee emp : employees) {
                                if (emp.getName() != null && emp.getName().trim().equalsIgnoreCase(name.trim())) {
                                    trade = emp.getTrade() != null ? emp.getTrade() : "N/A";
                                    break;
                                }
                            }

                            emailContent.append("<tr>");
                            emailContent.append("<td style='").append(tdStyle).append("'>").append(srNo++)
                                    .append("</td>");
                            emailContent.append("<td style='").append(tdStyle).append("'>").append(name.trim())
                                    .append("</td>");
                            emailContent.append("<td style='").append(tdStyle).append("'>").append(trade)
                                    .append("</td>");
                            emailContent.append("</tr>");
                        }
                    }
                    emailContent.append("</table>");
                } else {
                    emailContent.append("<p>- None</p>");
                }
                emailContent.append("<br>");
            }
        } else {
            // Fallback for Day 1
            emailContent.append("<h3>Day 1:</h3>");
            emailContent.append("<p><strong>Day 1 (").append(extractDate(day1Date)).append(") Summary:</strong><br>");
            emailContent.append("Present: ").append(summary.getDay1PresentCount()).append("<br>");
            emailContent.append("Absent: ").append(summary.getDay1AbsentCount()).append("</p>");

            emailContent.append("<h4>ABSENT STUDENTS TABLE FOR DAY 1 (").append(extractDate(day1Date))
                    .append("):</h4>");

            if (!summary.getDay1AbsentNames().isEmpty()) {
                emailContent.append("<table style='").append(tableStyle).append("'>");
                emailContent.append("<tr>");
                emailContent.append("<th style='").append(thStyle).append("'>Sr. No</th>");
                emailContent.append("<th style='").append(thStyle).append("'>Name</th>");
                emailContent.append("<th style='").append(thStyle).append("'>Trade/Department</th>");
                emailContent.append("</tr>");

                String[] names = summary.getDay1AbsentNames().split(", ");
                int srNo = 1;
                for (String name : names) {
                    if (!name.trim().isEmpty()) {
                        String trade = "N/A";
                        for (Employee emp : employees) {
                            if (emp.getName() != null && emp.getName().trim().equalsIgnoreCase(name.trim())) {
                                trade = emp.getTrade() != null ? emp.getTrade() : "N/A";
                                break;
                            }
                        }

                        emailContent.append("<tr>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(srNo++).append("</td>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(name.trim())
                                .append("</td>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(trade).append("</td>");
                        emailContent.append("</tr>");
                    }
                }
                emailContent.append("</table>");
            } else {
                emailContent.append("<p>No absent students recorded for this day</p>");
            }
            emailContent.append("<br>");

            // Fallback for Day 2
            emailContent.append("<h3>Day 2:</h3>");
            emailContent.append("<p><strong>Day 2 (").append(extractDate(day2Date)).append(") Summary:</strong><br>");
            emailContent.append("Present: ").append(summary.getDay2PresentCount()).append("<br>");
            emailContent.append("Absent: ").append(summary.getDay2AbsentCount()).append("</p>");

            emailContent.append("<h4>ABSENT STUDENTS TABLE FOR DAY 2 (").append(extractDate(day2Date))
                    .append("):</h4>");

            if (!summary.getDay2AbsentNames().isEmpty()) {
                emailContent.append("<table style='").append(tableStyle).append("'>");
                emailContent.append("<tr>");
                emailContent.append("<th style='").append(thStyle).append("'>Sr. No</th>");
                emailContent.append("<th style='").append(thStyle).append("'>Name</th>");
                emailContent.append("<th style='").append(thStyle).append("'>Trade/Department</th>");
                emailContent.append("</tr>");

                String[] names = summary.getDay2AbsentNames().split(", ");
                int srNo = 1;
                for (String name : names) {
                    if (!name.trim().isEmpty()) {
                        String trade = "N/A";
                        for (Employee emp : employees) {
                            if (emp.getName() != null && emp.getName().trim().equalsIgnoreCase(name.trim())) {
                                trade = emp.getTrade() != null ? emp.getTrade() : "N/A";
                                break;
                            }
                        }

                        emailContent.append("<tr>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(srNo++).append("</td>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(name.trim())
                                .append("</td>");
                        emailContent.append("<td style='").append(tdStyle).append("'>").append(trade).append("</td>");
                        emailContent.append("</tr>");
                    }
                }
                emailContent.append("</table>");
            } else {
                emailContent.append("<p>No absent students recorded for this day</p>");
            }
        }

        // Add NOT OK classification section
        emailContent.append("<h3>ATTENDANCE STATUS CLASSIFICATION:</h3>");
        
        // Group employees by attendance status
        java.util.List<Employee> notOkEmployees = new java.util.ArrayList<>();
        
        for (Employee emp : employees) {
            String day1Attendance = emp.getDay1Attendance();
            String day2Attendance = emp.getDay2Attendance();
            
            // Determine if employee is NOT OK (absent on one or both days)
            boolean isDay1Present = day1Attendance != null && 
                (day1Attendance.toUpperCase().contains("P") || day1Attendance.toUpperCase().contains("PRESENT"));
            boolean isDay2Present = day2Attendance != null && 
                (day2Attendance.toUpperCase().contains("P") || day2Attendance.toUpperCase().contains("PRESENT"));
            
            if (!(isDay1Present && isDay2Present)) {  // If not present on both days, then NOT OK
                notOkEmployees.add(emp);
            }
        }
        
        // NOT OK Section - Employees absent on one or both days
        emailContent.append("<h4>NOT OK (Absent on One or Both Days):</h4>");
        if (!notOkEmployees.isEmpty()) {
            emailContent.append("<table style='").append(tableStyle).append("'>");
            emailContent.append("<tr>");
            emailContent.append("<th style='").append(thStyle).append("'>Sr. No</th>");
            emailContent.append("<th style='").append(thStyle).append("'>Name</th>");
            emailContent.append("<th style='").append(thStyle).append("'>P No</th>");
            emailContent.append("<th style='").append(thStyle).append("'>Trade/Department</th>");
            emailContent.append("<th style='").append(thStyle).append("'>Day 1 Status</th>");
            emailContent.append("<th style='").append(thStyle).append("'>Day 2 Status</th>");
            emailContent.append("</tr>");
            
            int srNo = 1;
            for (Employee emp : notOkEmployees) {
                String day1Attendance = emp.getDay1Attendance();
                String day2Attendance = emp.getDay2Attendance();
                
                boolean isDay1Present = day1Attendance != null && 
                    (day1Attendance.toUpperCase().contains("P") || day1Attendance.toUpperCase().contains("PRESENT"));
                boolean isDay2Present = day2Attendance != null && 
                    (day2Attendance.toUpperCase().contains("P") || day2Attendance.toUpperCase().contains("PRESENT"));
                
                emailContent.append("<tr>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(srNo++).append("</td>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(emp.getName() != null ? emp.getName() : "N/A").append("</td>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(emp.getPNo() != null ? emp.getPNo() : "N/A").append("</td>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(emp.getTrade() != null ? emp.getTrade() : "N/A").append("</td>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(day1Attendance != null ? day1Attendance : "N/A").append("</td>");
                emailContent.append("<td style='").append(tdStyle).append("'>").append(day2Attendance != null ? day2Attendance : "N/A").append("</td>");
                emailContent.append("</tr>");
            }
            emailContent.append("</table>");
        } else {
            emailContent.append("<p>No employees marked as NOT OK (absent on one or both days)</p>");
        }
        
        emailContent.append("</body></html>");
        return emailContent.toString();
    }

    // Helper method to format date from various formats (e.g., "16-Dec" ->
    // "16-Dec-2025")
    private String formatDate(String dateHeader) {
        if (dateHeader == null || dateHeader.isEmpty()) {
            return "N/A";
        }

        // If the date is already in the format like "16-Dec-2025", return as is
        if (dateHeader.matches("\\d{1,2}-[A-Za-z]{3}-\\d{4}")) {
            return dateHeader;
        }

        // If the date is in the format like "16-Dec", add the year
        if (dateHeader.matches("\\d{1,2}-[A-Za-z]{3}")) {
            // Add a reasonable default year (current year or based on context)
            return dateHeader + "-" + java.time.LocalDate.now().getYear();
        }

        // Extract date from format like "Day 1 (05/01/2026)" or similar
        if (dateHeader.contains("(") && dateHeader.contains(")")) {
            int start = dateHeader.indexOf('(');
            int end = dateHeader.indexOf(')');
            return dateHeader.substring(start + 1, end);
        }

        return dateHeader; // Return as is if not in expected format
    }

    // Helper method to extract date from column header like "Day 1 (05/01/2026)"
    private String extractDate(String dateHeader) {
        if (dateHeader == null || dateHeader.isEmpty()) {
            return "N/A";
        }

        // Extract date from format like "Day 1 (05/01/2026)" or similar
        if (dateHeader.contains("(") && dateHeader.contains(")")) {
            int start = dateHeader.indexOf('(');
            int end = dateHeader.indexOf(')');
            return dateHeader.substring(start + 1, end);
        }

        return dateHeader; // Return as is if not in expected format
    }
}