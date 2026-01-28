import java.util.Scanner;

import model.Config;

/**
 * Utility to help update the email configuration manually
 */
public class ManualConfigUpdater {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        
        System.out.println("=== Gmail Configuration Updater ===");
        System.out.println("Current Configuration:");
        System.out.println("Email: " + Config.getSenderEmail());
        System.out.println("Password length: " + Config.getSenderAppPassword().length());
        System.out.println();
        
        System.out.println("To fix the Gmail authentication issue:");
        System.out.println("1. Go to your Google Account settings");
        System.out.println("2. Enable 2-factor authentication if not already enabled");
        System.out.println("3. Generate a new APP PASSWORD (not your regular password)");
        System.out.println("4. The app password will be 16 characters with spaces");
        System.out.println();
        
        System.out.print("Enter your NEW 16-character Gmail App Password: ");
        String newPassword = scanner.nextLine().trim();
        
        if (newPassword.length() != 16 && !newPassword.contains(" ")) {
            System.out.println("Warning: App password should be 16 characters. Are you sure this is correct?");
            System.out.print("Continue anyway? (y/n): ");
            String response = scanner.nextLine().trim();
            if (!response.toLowerCase().startsWith("y")) {
                System.out.println("Operation cancelled.");
                scanner.close();
                return;
            }
        }
        
        System.out.println();
        System.out.println("To update Config.java with the new password:");
        System.out.println("1. Open: src/main/java/model/Config.java");
        System.out.println("2. Find the line with: private static String SENDER_APP_PASSWORD");
        System.out.println("3. Replace the current password with: \"" + newPassword + "\"");
        System.out.println("4. Save the file");
        System.out.println();
        System.out.println("After updating, run: mvn clean compile");
        System.out.println("Then test with: mvn exec:java@test-email");
        
        scanner.close();
    }
}