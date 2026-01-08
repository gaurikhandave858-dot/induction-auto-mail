import javax.swing.SwingUtilities;

import view.AttendanceEmailFrame;

public class Main {
    public static void main(String[] args) {
        // Set look and feel to system default
        try {
            javax.swing.UIManager.setLookAndFeel(javax.swing.UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        // Launch the GUI application
        SwingUtilities.invokeLater(() -> {
            new AttendanceEmailFrame().setVisible(true);
        });
    }
}