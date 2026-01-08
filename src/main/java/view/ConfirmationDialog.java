package view;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.Frame;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.SwingConstants;

public class ConfirmationDialog extends JDialog {
    private boolean confirmed = false;
    
    public ConfirmationDialog(Frame parent, String message) {
        super(parent, "Confirmation", true); // Modal dialog
        initializeComponents(message);
        setupLayout();
        setupEventHandlers();
        setSize(500, 250);
        setLocationRelativeTo(parent);
    }
    
    private void initializeComponents(String message) {
        setLayout(new BorderLayout());
        
        // Message panel with increased padding
        JPanel messagePanel = new JPanel(new BorderLayout());
        JLabel messageLabel = new JLabel(message, SwingConstants.CENTER);
        messageLabel.setFont(messageLabel.getFont().deriveFont(16.0f)); // Increase font size
        messageLabel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20)); // Add padding around the text
        messagePanel.add(messageLabel, BorderLayout.CENTER);
        
        // Button panel
        JPanel buttonPanel = new JPanel(new FlowLayout());
        JButton btnYes = new JButton("Yes");
        JButton btnNo = new JButton("No");
        
        // Increase font size for better visibility
        btnYes.setFont(btnYes.getFont().deriveFont(14.0f));
        btnNo.setFont(btnNo.getFont().deriveFont(14.0f));
        
        buttonPanel.add(btnYes);
        buttonPanel.add(btnNo);
        
        add(messagePanel, BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);
        
        // Event handlers
        btnYes.addActionListener(e -> {
            confirmed = true;
            dispose();
        });
        
        btnNo.addActionListener(e -> {
            confirmed = false;
            dispose();
        });
        
        // Allow closing the dialog with ESC key
        getRootPane().setDefaultButton(btnNo);
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
    }
    
    private void setupLayout() {
        // Layout is handled in initializeComponents
    }
    
    private void setupEventHandlers() {
        // Event handlers are handled in initializeComponents
    }
    
    public boolean isConfirmed() {
        return confirmed;
    }
}