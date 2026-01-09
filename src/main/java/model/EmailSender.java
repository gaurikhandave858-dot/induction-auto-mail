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
            
            // Create session with authentication
            Session session = Session.getInstance(props, new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(email, password);
                }
            });
            
            // Create message
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(email));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject(subject);
            message.setText(body);
            
            // Send message
            Transport.send(message);
            
            System.out.println("Email sent successfully to: " + recipientEmail);
            return true;
        } catch (MessagingException e) {
            e.printStackTrace();
            System.err.println("Failed to send email: " + e.getMessage());
            return false;
        }
    }
    
    public String generateAttendanceSummaryEmail(AttendanceSummary summary) {
        StringBuilder emailContent = new StringBuilder();
        
        emailContent.append("Subject: Daily Attendance Summary Report\n\n");
        emailContent.append("Dear Team,\n\n");
        emailContent.append("Please find the attendance summary for the induction program:\n\n");
        
        // Day 1 Summary
        emailContent.append("Day 1 Summary:\n");
        emailContent.append("- Present: ").append(summary.getDay1PresentCount()).append("\n");
        emailContent.append("- Absent: ").append(summary.getDay1AbsentCount()).append("\n");
        if (!summary.getDay1AbsentNames().isEmpty()) {
            emailContent.append("- Absent Names: ").append(summary.getDay1AbsentNames()).append("\n");
        }
        emailContent.append("\n");
        
        // Day 2 Summary
        emailContent.append("Day 2 Summary:\n");
        emailContent.append("- Present: ").append(summary.getDay2PresentCount()).append("\n");
        emailContent.append("- Absent: ").append(summary.getDay2AbsentCount()).append("\n");
        if (!summary.getDay2AbsentNames().isEmpty()) {
            emailContent.append("- Absent Names: ").append(summary.getDay2AbsentNames()).append("\n");
        }
        emailContent.append("\n");
        
        emailContent.append("Best regards,\n");
        emailContent.append("Automated Attendance System");
        
        return emailContent.toString();
    }
    
    // Method to generate email with the exact format as specified in SRS
    public String generateAttendanceSummaryEmailWithDates(AttendanceSummary summary, String day1Date, String day2Date) {
        StringBuilder emailContent = new StringBuilder();
        
        emailContent.append("EMAIL SUMMARY CONTENT\n\n");
        emailContent.append("Report Generated On: ").append(java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"))).append("\n\n");
        
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
                
                // Format the date properly (extract date from format like "16-Dec")
                String formattedDate = formatDate(dateHeader);
                
                // Add day header (e.g., "Day 1:")
                emailContent.append("Day ").append(i + 1).append(":\n\n");
                
                // Add summary for this day in SRS format
                emailContent.append("Day ").append(i + 1).append(" (").append(formattedDate).append(") Summary:\n");
                emailContent.append("Present: ").append(presentCount).append("\n");
                emailContent.append("Absent: ").append(absentCount).append("\n");
                
                // Add absent employee list header for this day
                emailContent.append("Day ").append(i + 1).append(" (").append(formattedDate).append(") Absent Employee List:\n\n");
                
                if (!absentNames.isEmpty()) {
                    for (String name : absentNames) {
                        if (name != null && !name.trim().isEmpty()) {
                            emailContent.append("- ").append(name.trim()).append("\n");
                        }
                    }
                } else {
                    emailContent.append("- None\n");
                }
                emailContent.append("\n");
            }
        } else {
            // Fallback to original logic for backward compatibility
            
            // Day 1 header
            emailContent.append("Day 1:\n\n");
            
            // Day 1 Summary with counts in SRS format
            emailContent.append("Day 1 (").append(extractDate(day1Date)).append(") Summary:\n");
            emailContent.append("Present: ").append(summary.getDay1PresentCount()).append("\n");
            emailContent.append("Absent: ").append(summary.getDay1AbsentCount()).append("\n");
            
            // Day 1 Absent Employee List header
            emailContent.append("Day 1 (").append(extractDate(day1Date)).append(") Absent Employee List:\n\n");
            if (!summary.getDay1AbsentNames().isEmpty()) {
                // Format absent names with each on a new line with dash
                String[] names = summary.getDay1AbsentNames().split(", ");
                for (int i = 0; i < names.length; i++) {
                    emailContent.append("- ").append(names[i].trim()).append("\n");
                }
            } else {
                emailContent.append("- None\n");
            }
            emailContent.append("\n");
            
            // Day 2 Summary with counts in SRS format
            emailContent.append("Day 2 (").append(extractDate(day2Date)).append(") Summary:\n");
            emailContent.append("Present: ").append(summary.getDay2PresentCount()).append("\n");
            emailContent.append("Absent: ").append(summary.getDay2AbsentCount()).append("\n");
            
            // Day 2 Absent Employee List header
            emailContent.append("Day 2 (").append(extractDate(day2Date)).append(") Absent Employee List:\n\n");
            if (!summary.getDay2AbsentNames().isEmpty()) {
                // Format absent names with each on a new line with dash
                String[] names = summary.getDay2AbsentNames().split(", ");
                for (int i = 0; i < names.length; i++) {
                    emailContent.append("- ").append(names[i].trim()).append("\n");
                }
            } else {
                emailContent.append("- None\n");
            }
            emailContent.append("\n");
        }
        
        return emailContent.toString();
    }
    
    // Helper method to format date from various formats (e.g., "16-Dec" -> "16-Dec-2025")
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