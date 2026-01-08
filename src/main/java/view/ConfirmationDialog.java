package view;

import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.Frame;

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
        setSize(300, 150);
        setLocationRelativeTo(parent);
    }
    
    private void initializeComponents(String message) {
        setLayout(new BorderLayout());
        
        // Message panel
        JPanel messagePanel = new JPanel();
        messagePanel.add(new JLabel(message, SwingConstants.CENTER));
        
        // Button panel
        JPanel buttonPanel = new JPanel(new FlowLayout());
        JButton btnYes = new JButton("Yes");
        JButton btnNo = new JButton("No");
        
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