package view;

import controller.AttendanceController;
import model.Employee;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;

import view.ConfirmationDialog;

public class AttendanceEmailFrame extends JFrame {
    private JTextField txtAttendanceFile;
    private JTextField txtRecipientEmail;
    private JButton btnBrowse;
    private JButton btnProcess;
    private JButton btnSendEmail;
    private JTextArea txtResult;
    private JScrollPane scrollPane;
    private AttendanceController controller;
    private java.util.List<model.Employee> employees;
    private JComboBox<String> cmbAttendanceStatus;
    private JComboBox<String> cmbDayFilter;
    private JComboBox<String> cmbAbsenteeismType;
    
    public AttendanceEmailFrame() {
        initializeComponents();
        setupLayout();
        setupEventHandlers();
        controller = new AttendanceController();
    }
    
    private void initializeComponents() {
        setTitle("Automated Attendance Tracking and Email Reporting System");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 800);
        setLocationRelativeTo(null); // Center the window
        
        txtAttendanceFile = new JTextField(50);
        txtAttendanceFile.setFont(new Font("Arial", Font.BOLD, 16));
        txtRecipientEmail = new JTextField(50);
        txtRecipientEmail.setFont(new Font("Arial", Font.BOLD, 16));
        btnBrowse = new JButton("Browse Attendance File");
        btnBrowse.setFont(new Font("Arial", Font.BOLD, 16));
        btnBrowse.setPreferredSize(new Dimension(200, 35));
        btnProcess = new JButton("Process Summary Email");
        btnProcess.setFont(new Font("Arial", Font.BOLD, 16));
        btnProcess.setPreferredSize(new Dimension(200, 35));
        btnSendEmail = new JButton("Send Email");
        btnSendEmail.setFont(new Font("Arial", Font.BOLD, 16));
        btnSendEmail.setPreferredSize(new Dimension(140, 35));
        txtResult = new JTextArea(20, 70);
        txtResult.setFont(new Font("Arial", Font.PLAIN, 14));
        txtResult.setEditable(false);
        scrollPane = new JScrollPane(txtResult);
        
        // Make buttons initially disabled until file is selected
        btnProcess.setEnabled(false);
        btnSendEmail.setEnabled(false);
    }
    
    private void setupLayout() {
        setLayout(new BorderLayout());
        
        // Top panel for file selection
        JPanel topPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        
        gbc.gridx = 0; gbc.gridy = 0;
        JLabel lblAttendanceFile = new JLabel("Attendance File:");
        lblAttendanceFile.setFont(new Font("Arial", Font.BOLD, 16));
        topPanel.add(lblAttendanceFile, gbc);
        
        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL; // Make the text field fill available space
        gbc.weightx = 1.0; // Allow the text field to expand
        topPanel.add(txtAttendanceFile, gbc);
        gbc.fill = GridBagConstraints.NONE; // Reset fill
        gbc.weightx = 0.0; // Reset weight
        
        gbc.gridx = 2;
        topPanel.add(btnBrowse, gbc);
        
        // Middle panel for email recipient
        JPanel middlePanel = new JPanel(new GridBagLayout());
        gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        
        gbc.gridx = 0; gbc.gridy = 0;
        JLabel lblRecipientEmail = new JLabel("Recipient Email:");
        lblRecipientEmail.setFont(new Font("Arial", Font.BOLD, 16));
        middlePanel.add(lblRecipientEmail, gbc);
        
        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL; // Make the text field fill available space
        gbc.weightx = 1.0; // Allow the text field to expand
        middlePanel.add(txtRecipientEmail, gbc);
        gbc.fill = GridBagConstraints.NONE; // Reset fill
        gbc.weightx = 0.0; // Reset weight
        
        gbc.gridx = 2;
        middlePanel.add(btnProcess, gbc);
        
        gbc.gridx = 3;
        middlePanel.add(btnSendEmail, gbc);
        
        // Add components to frame
        add(topPanel, BorderLayout.NORTH);
        add(middlePanel, BorderLayout.CENTER);
        
        // Bottom panel for filters
        JPanel filterPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc2 = new GridBagConstraints();
        gbc2.insets = new Insets(5, 5, 5, 5);
        
        // Attendance Status Filter
        gbc2.gridx = 0; gbc2.gridy = 0;
        JLabel lblAttendanceFilter = new JLabel("Attendance Status Filter:");
        lblAttendanceFilter.setFont(new Font("Arial", Font.BOLD, 14));
        filterPanel.add(lblAttendanceFilter, gbc2);
        
        String[] statusOptions = {"ALL", "PRESENT", "ABSENT"};
        this.cmbAttendanceStatus = new JComboBox<>(statusOptions);
        this.cmbAttendanceStatus.setFont(new Font("Arial", Font.PLAIN, 14));
        this.cmbAttendanceStatus.setPreferredSize(new Dimension(120, 30));
        gbc2.gridx = 1;
        filterPanel.add(this.cmbAttendanceStatus, gbc2);
        
        String[] dayOptions = {"Day 1", "Day 2"};
        this.cmbDayFilter = new JComboBox<>(dayOptions);
        this.cmbDayFilter.setFont(new Font("Arial", Font.PLAIN, 14));
        this.cmbDayFilter.setPreferredSize(new Dimension(100, 30));
        gbc2.gridx = 2;
        filterPanel.add(this.cmbDayFilter, gbc2);
        
        JButton btnApplyStatusFilter = new JButton("Apply Status Filter");
        btnApplyStatusFilter.setFont(new Font("Arial", Font.BOLD, 14));
        btnApplyStatusFilter.setPreferredSize(new Dimension(150, 30));
        gbc2.gridx = 3;
        filterPanel.add(btnApplyStatusFilter, gbc2);
        
        // Absenteeism Filter
        gbc2.gridx = 4; gbc2.gridy = 0;
        JLabel lblAbsenteeismFilter = new JLabel("Absenteeism Filter:");
        lblAbsenteeismFilter.setFont(new Font("Arial", Font.BOLD, 14));
        filterPanel.add(lblAbsenteeismFilter, gbc2);
        
        String[] absenteeismOptions = {"ALL", "Absent on Day 1 Only", "Absent on Day 2 Only", "Absent on Both Days", "Absent on Any Day"};
        this.cmbAbsenteeismType = new JComboBox<>(absenteeismOptions);
        this.cmbAbsenteeismType.setFont(new Font("Arial", Font.PLAIN, 14));
        this.cmbAbsenteeismType.setPreferredSize(new Dimension(180, 30));
        gbc2.gridx = 5;
        filterPanel.add(this.cmbAbsenteeismType, gbc2);
        
        JButton btnApplyAbsenteeismFilter = new JButton("Apply Absenteeism Filter");
        btnApplyAbsenteeismFilter.setFont(new Font("Arial", Font.BOLD, 14));
        btnApplyAbsenteeismFilter.setPreferredSize(new Dimension(180, 30));
        gbc2.gridx = 6;
        filterPanel.add(btnApplyAbsenteeismFilter, gbc2);
        
        // Add action listeners for filter buttons
        btnApplyStatusFilter.addActionListener(e -> applyAttendanceStatusFilter(
            this.cmbAttendanceStatus.getSelectedItem().toString(), 
            this.cmbDayFilter.getSelectedItem().toString()));
        
        btnApplyAbsenteeismFilter.addActionListener(e -> applyAbsenteeismFilter(
            this.cmbAbsenteeismType.getSelectedItem().toString()));
        
        add(filterPanel, BorderLayout.SOUTH);
        add(scrollPane, BorderLayout.PAGE_END);
    }
    
    private void setupEventHandlers() {
        btnBrowse.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                browseFile();
            }
        });
        
        btnProcess.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                processAttendance();
            }
        });
        
        btnSendEmail.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                sendEmailConfirmation();
            }
        });
    }
    
    private void browseFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files (.xls, .xlsx)", "xls", "xlsx"));
        
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            txtAttendanceFile.setText(selectedFile.getAbsolutePath());
            btnProcess.setEnabled(true);
            txtResult.append("Selected file: " + selectedFile.getName() + "\n");
        }
    }
    
    private void processAttendance() {
        String filePath = txtAttendanceFile.getText().trim();
        if (filePath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please select an attendance file first.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        try {
            txtResult.append("Processing attendance data...\n");
            
            // Process the attendance file
            this.employees = controller.processAttendanceFile(filePath);
            String emailContent = controller.generateEmailSummary(employees);
            
            txtResult.append("Attendance processed successfully!\n");
            txtResult.append("Day 1 Present: " + controller.getDay1PresentCount() + "\n");
            txtResult.append("Day 1 Absent: " + controller.getDay1AbsentCount() + "\n");
            txtResult.append("Day 2 Present: " + controller.getDay2PresentCount() + "\n");
            txtResult.append("Day 2 Absent: " + controller.getDay2AbsentCount() + "\n");
            txtResult.append("\nEmail Content Preview:\n");
            txtResult.append(emailContent + "\n");
            
            btnSendEmail.setEnabled(true);
            
        } catch (Exception ex) {
            txtResult.append("Error processing attendance: " + ex.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "Error processing attendance file: " + ex.getMessage(), 
                                        "Error", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }
    }
    
    private void sendEmailConfirmation() {
        String recipientEmail = txtRecipientEmail.getText().trim();
        if (recipientEmail.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please enter recipient email address.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        // Show confirmation dialog
        ConfirmationDialog confirmDialog = new ConfirmationDialog(this, 
            "Are you sure you want to send the email to " + recipientEmail + "?");
        confirmDialog.setVisible(true);
        
        if (confirmDialog.isConfirmed()) {
            sendEmail();
        }
    }
    
    private void sendEmail() {
        String recipientEmail = txtRecipientEmail.getText().trim();
        if (recipientEmail.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please enter recipient email address.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        txtResult.append("Sending email to: " + recipientEmail + "\n");
        
        boolean success = controller.sendEmail(recipientEmail);
        if (success) {
            txtResult.append("Email sent successfully!\n");
            
            // Also send a copy to self
            boolean selfSuccess = controller.sendEmailToSelf();
            if (selfSuccess) {
                txtResult.append("Self email sent successfully!\n");
            } else {
                txtResult.append("Failed to send self email.\n");
            }
            

        } else {
            txtResult.append("Failed to send email.\n");
            JOptionPane.showMessageDialog(this, "Failed to send email. Please check your email settings.", 
                                        "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void applyAttendanceStatusFilter(String status, String day) {
        if (employees == null || employees.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please process an attendance file first.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        try {
            model.AttendanceFilter.AttendanceStatus attendanceStatus;
            
            switch (status.toUpperCase()) {
                case "PRESENT":
                    attendanceStatus = model.AttendanceFilter.AttendanceStatus.PRESENT;
                    break;
                case "ABSENT":
                    attendanceStatus = model.AttendanceFilter.AttendanceStatus.ABSENT;
                    break;
                case "ALL":
                default:
                    attendanceStatus = model.AttendanceFilter.AttendanceStatus.ALL;
                    break;
            }
            
            String dayValue = day.contains("1") ? "day1" : "day2";
            java.util.List<model.Employee> filteredEmployees = controller.filterEmployeesByStatus(attendanceStatus, dayValue);
            
            txtResult.append("\nFilter applied: " + status + " employees on " + day + "\n");
            txtResult.append("Number of employees: " + filteredEmployees.size() + "\n");
            
            if (!filteredEmployees.isEmpty()) {
                txtResult.append("Employee Names: " + String.join(", ", 
                    filteredEmployees.stream().map(model.Employee::getName).toArray(String[]::new)) + "\n");
            }
            
        } catch (Exception ex) {
            txtResult.append("Error applying attendance status filter: " + ex.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "Error applying filter: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void applyAbsenteeismFilter(String absenteeismType) {
        if (employees == null || employees.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please process an attendance file first.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        try {
            model.AttendanceFilter.AbsenteeismType absType;
            
            switch (absenteeismType) {
                case "Absent on Day 1 Only":
                    absType = model.AttendanceFilter.AbsenteeismType.ABSENT_ON_DAY1;
                    break;
                case "Absent on Day 2 Only":
                    absType = model.AttendanceFilter.AbsenteeismType.ABSENT_ON_DAY2;
                    break;
                case "Absent on Both Days":
                    absType = model.AttendanceFilter.AbsenteeismType.ABSENT_ON_BOTH_DAYS;
                    break;
                case "Absent on Any Day":
                    absType = model.AttendanceFilter.AbsenteeismType.ABSENT_ON_ANY_DAY;
                    break;
                case "ALL":
                default:
                    absType = model.AttendanceFilter.AbsenteeismType.ALL;
                    break;
            }
            
            java.util.List<model.Employee> filteredEmployees = controller.filterEmployeesByAbsenteeism(absType);
            
            txtResult.append("\nFilter applied: " + absenteeismType + "\n");
            txtResult.append("Number of employees: " + filteredEmployees.size() + "\n");
            
            if (!filteredEmployees.isEmpty()) {
                txtResult.append("Employee Names: " + String.join(", ", 
                    filteredEmployees.stream().map(model.Employee::getName).toArray(String[]::new)) + "\n");
            }
            
        } catch (Exception ex) {
            txtResult.append("Error applying absenteeism filter: " + ex.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "Error applying filter: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}