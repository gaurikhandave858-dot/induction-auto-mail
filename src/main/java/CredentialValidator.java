import model.Config;

public class CredentialValidator {
    public static void main(String[] args) {
        System.out.println("=== Gmail Credential Validator ===");
        System.out.println("Checking your email configuration...");
        System.out.println();
        
        System.out.println("Email: " + Config.getSenderEmail());
        System.out.println("Password length: " + Config.getSenderAppPassword().length());
        System.out.println("SMTP Host: " + Config.getSmtpHost());
        System.out.println("SMTP Port: " + Config.getSmtpPort());
        System.out.println();
        
        System.out.println("Validating credentials...");
        boolean isValid = Config.validateCredentials();
        
        if (isValid) {
            System.out.println("✓ Credentials are valid! Email sending should work.");
        } else {
            System.out.println("✗ Credentials are invalid. Please check:");
            System.out.println("  1. Your Gmail account has 2-factor authentication enabled");
            System.out.println("  2. You generated an app password specifically for 'Mail'");
            System.out.println("  3. The app password is exactly 16 characters");
            System.out.println("  4. The email address matches your Gmail account");
            System.out.println("  5. You're using the app password, not your regular password");
        }
        
        System.out.println();
        System.out.println("=== Troubleshooting Steps ===");
        System.out.println("If validation fails:");
        System.out.println("1. Go to Google Account Settings");
        System.out.println("2. Enable 2-factor authentication");
        System.out.println("3. Generate a new app password for 'Mail'");
        System.out.println("4. Update Config.java with the new password");
        System.out.println("5. Rebuild and test again");
    }
}