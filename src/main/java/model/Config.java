package model;

import java.util.Arrays;
import java.util.List;

/**
 * Configuration class to store email settings and receiver information
 * for the Automated Attendance System.
 */
public class Config {
    
    // Email sender configuration
    private static String SENDER_EMAIL = "gaurikhandave858@gmail.com";
    private static String SENDER_APP_PASSWORD = "mnqjxlqixiochbzm";

    
    // SMTP configuration
    private static String SMTP_HOST = "smtp.gmail.com";
    
    private static int SMTP_PORT = 587;
    private static boolean SMTP_TLS_ENABLED = true;
    
    // Email receivers - list of email addresses to send attendance summaries to
    private static List<String> RECEIVER_EMAILS = Arrays.asList(
        "gaurikhandave858@gmail.com");  // Update with actual recipient emails
    
    // Setter method to update receiver emails dynamically
    public static void setReceiverEmails(List<String> emails) {
        RECEIVER_EMAILS = emails;
    }
    
    public static void setSmtpHost(String smtpHost) {
        Config.SMTP_HOST = smtpHost;
    }
    
    public static void setSmtpPort(int smtpPort) {
        Config.SMTP_PORT = smtpPort;
    }
    
    public static void setSenderEmail(String senderEmail) {
        Config.SENDER_EMAIL = senderEmail;
    }
    
    public static void setSenderAppPassword(String senderAppPassword) {
        Config.SENDER_APP_PASSWORD = senderAppPassword;
    }
    
    public static void setSmtpTlsEnabled(boolean smtpTlsEnabled) {
        Config.SMTP_TLS_ENABLED = smtpTlsEnabled;
    }
    
    // Email subject
    private static final String EMAIL_SUBJECT = "Attendance Summary Report";
    
    // Getter methods
    public static String getSenderEmail() {
        return SENDER_EMAIL;
    }
    
    public static String getSenderAppPassword() {
        return SENDER_APP_PASSWORD;
    }
    
    public static String getSmtpHost() {
        return SMTP_HOST;
    }
    
    public static int getSmtpPort() {
        return SMTP_PORT;
    }
    
    public static boolean isSmtpTlsEnabled() {
        return SMTP_TLS_ENABLED;
    }
    
    public static List<String> getReceiverEmails() {
        return RECEIVER_EMAILS;
    }
    
    public static String getEmailSubject() {
        return EMAIL_SUBJECT;
    }
    
    /**
     * Method to validate if configuration is properly set
     * @return true if all required configurations are set, false otherwise
     */
    public static boolean isConfigValid() {
        return SENDER_EMAIL != null && !SENDER_EMAIL.isEmpty() &&
               SENDER_APP_PASSWORD != null && !SENDER_APP_PASSWORD.isEmpty() &&
               SMTP_HOST != null && !SMTP_HOST.isEmpty() &&
               RECEIVER_EMAILS != null && !RECEIVER_EMAILS.isEmpty();
    }

    // Method to validate credentials by attempting to create a mail session
    public static boolean validateCredentials() {
        try {
            java.util.Properties props = new java.util.Properties();
            props.put("mail.smtp.host", SMTP_HOST);
            props.put("mail.smtp.port", String.valueOf(SMTP_PORT));
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.starttls.enable", "true");
            
            javax.mail.Session session = javax.mail.Session.getInstance(props,
                new javax.mail.Authenticator() {
                    @Override
                    protected javax.mail.PasswordAuthentication getPasswordAuthentication() {
                        return new javax.mail.PasswordAuthentication(SENDER_EMAIL, SENDER_APP_PASSWORD);
                    }
                });
            
            // Try to connect to the SMTP server
            javax.mail.Transport transport = session.getTransport("smtp");
            transport.connect(SMTP_HOST, SENDER_EMAIL, SENDER_APP_PASSWORD);
            transport.close();
            
            return true;
        } catch (Exception e) {
            System.out.println("Credential validation failed: " + e.getMessage());
            return false;
        }
    }
}