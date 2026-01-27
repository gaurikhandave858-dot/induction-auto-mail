import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import model.Config;

/**
 * Test class to verify email configuration
 */
public class EmailTest {
    public static void main(String[] args) {
        System.out.println("Testing email configuration from Config class...");
        System.out.println("SMTP Host: " + Config.getSmtpHost());
        System.out.println("SMTP Port: " + Config.getSmtpPort());
        System.out.println("Sender Email: " + Config.getSenderEmail());
        System.out.println("Sender Email Valid: " + (Config.getSenderEmail() != null && !Config.getSenderEmail().isEmpty()));
        System.out.println("Sender Password Valid: " + (Config.getSenderAppPassword() != null && !Config.getSenderAppPassword().isEmpty()));
        
        // Validate inputs
        if (Config.getSenderEmail().isEmpty() || Config.getSenderAppPassword().isEmpty()) {
            System.out.println("ERROR: Sender email and/or password not configured in Config class!");
            return;
        }
        
        String recipientEmail = Config.getSenderEmail(); // Send test to self
        
        // Test email sending
        try {
            // Setup mail server properties
            Properties props = new Properties();
            props.put("mail.smtp.host", Config.getSmtpHost());
            props.put("mail.smtp.port", String.valueOf(Config.getSmtpPort()));
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.starttls.enable", "true");
            props.put("mail.smtp.connectiontimeout", "10000");
            props.put("mail.smtp.timeout", "10000");
            props.put("mail.smtp.writetimeout", "10000");

            // Create session with authentication
            Session session = Session.getInstance(props, new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(Config.getSenderEmail(), Config.getSenderAppPassword());
                }
            });

            // Create message
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(Config.getSenderEmail()));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recipientEmail));
            message.setSubject("Test Email - Attendance System");
            message.setText("This is a test email from the attendance tracking system.\n\nConfiguration is working correctly!");

            // Send message
            System.out.println("Attempting to send test email...");
            Transport.send(message);

            System.out.println("SUCCESS: Test email sent successfully!");
            
        } catch (MessagingException e) {
            System.out.println("MessagingException occurred:");
            System.out.println("  Exception Type: " + e.getClass().getSimpleName());
            System.out.println("  Message: " + e.getMessage());
            
            // Check for specific error types
            Throwable cause = e.getCause();
            if (cause != null) {
                System.out.println("  Cause: " + cause.getMessage());
            }
            
            System.out.println("\nCommon causes for email failure:");
            System.out.println("  1. Invalid email address or password");
            System.out.println("  2. Gmail security settings blocking access");
            System.out.println("  3. Network connectivity issues");
            System.out.println("  4. App password not set up for Gmail");
            System.out.println("  5. Two-factor authentication required but not configured");
            System.out.println("\nFor Gmail, make sure:");
            System.out.println("  - You have enabled 2-factor authentication");
            System.out.println("  - You have generated an app password");
            System.out.println("  - You are using the app password, not your regular password");
            
            e.printStackTrace();
        } catch (Exception e) {
            System.out.println("General exception occurred:");
            System.out.println("  Exception Type: " + e.getClass().getSimpleName());
            System.out.println("  Message: " + e.getMessage());
            e.printStackTrace();
        }
    }
}