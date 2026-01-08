package model;

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
    
    // Method to generate email with the exact format as specified
    public String generateAttendanceSummaryEmailWithDates(AttendanceSummary summary, String day1Date, String day2Date) {
        StringBuilder emailContent = new StringBuilder();
        
        emailContent.append("EMAIL SUMMARY CONTENT\n\n");
        emailContent.append("Report Generated On: ").append(java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"))).append("\n\n");
        
        // Day 1 header
        emailContent.append("Day 1:\n");
        
        // Day 1 Summary with counts
        emailContent.append("Day 1 (").append(extractDate(day1Date)).append(") Summary:\n");
        emailContent.append("Present: ").append(summary.getDay1PresentCount()).append("\n");
        emailContent.append("Absent: ").append(summary.getDay1AbsentCount()).append("\n\n");
        
        // Day 1 Absent Employee List
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
        
        // Day 2 Summary with counts
        emailContent.append("Day 2 (").append(extractDate(day2Date)).append(") Summary:\n");
        emailContent.append("Present: ").append(summary.getDay2PresentCount()).append("\n");
        emailContent.append("Absent: ").append(summary.getDay2AbsentCount()).append("\n\n");
        
        // Day 2 Absent Employee List
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
        
        return emailContent.toString();
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