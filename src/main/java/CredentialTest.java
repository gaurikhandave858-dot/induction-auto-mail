import model.Config;

public class CredentialTest {
    public static void main(String[] args) {
        System.out.println("=== Email Credential Verification ===");
        System.out.println("Sender Email: " + Config.getSenderEmail());
        System.out.println("App Password Length: " + Config.getSenderAppPassword().length());
        System.out.println("App Password (masked): " + maskPassword(Config.getSenderAppPassword()));
        System.out.println("SMTP Host: " + Config.getSmtpHost());
        System.out.println("SMTP Port: " + Config.getSmtpPort());
        System.out.println("TLS Enabled: " + Config.isSmtpTlsEnabled());
        System.out.println("=====================================");
        
        // Validate configuration
        if (Config.isConfigValid()) {
            System.out.println("✓ Configuration appears valid");
        } else {
            System.out.println("✗ Configuration is invalid");
        }
        
        System.out.println("\nNext steps:");
        System.out.println("1. Verify 2-factor authentication is enabled on your Gmail account");
        System.out.println("2. Generate a new app password from Google Account settings");
        System.out.println("3. Update Config.java with the new app password");
        System.out.println("4. Run 'mvn exec:java@test-email' to test");
    }
    
    private static String maskPassword(String password) {
        if (password == null || password.isEmpty()) {
            return "[EMPTY]";
        }
        return "*".repeat(Math.min(password.length(), 16));
    }
}