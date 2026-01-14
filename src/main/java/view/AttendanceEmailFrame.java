package view;

import controller.AttendanceController;
import model.Employee;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.BorderFactory;
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
    private JEditorPane previewPane;
    private JTable emailTable;
    private JScrollPane scrollPane;
    private JScrollPane tableScrollPane;
    private JTabbedPane resultTabbedPane;
    private AttendanceController controller;
    private java.util.List<model.Employee> employees;
    private JComboBox<String> cmbAttendanceStatus;
    private JComboBox<String> cmbDayFilter;
    private JComboBox<String> cmbAbsenteeismType;
    
    // Define color scheme
    private final Color primaryColor = new Color(32, 102, 148);  // Dark blue
    private final Color secondaryColor = new Color(0, 150, 199); // Light blue
    private final Color accentColor = new Color(76, 175, 80);    // Green
    private final Color backgroundColor = new Color(245, 245, 245); // Light gray
    private final Color textColor = new Color(33, 33, 33);      // Dark gray

    public AttendanceEmailFrame() {
        initializeComponents();
        setupLayout();
        setupEventHandlers();
        controller = new AttendanceController();
    }

    private void initializeComponents() {
        setTitle("Automated Attendance Tracking and Email Reporting System");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 900);
        setLocationRelativeTo(null); // Center the window
        getContentPane().setBackground(backgroundColor);

        txtAttendanceFile = new JTextField(50);
        txtAttendanceFile.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        txtAttendanceFile.setBackground(Color.WHITE);
        txtAttendanceFile.setForeground(textColor);
        txtAttendanceFile.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(primaryColor, 2),
            BorderFactory.createEmptyBorder(5, 10, 5, 10)
        ));
        
        txtRecipientEmail = new JTextField(50);
        txtRecipientEmail.setFont(new Font("Segoe UI", Font.PLAIN, 16));
        txtRecipientEmail.setBackground(Color.WHITE);
        txtRecipientEmail.setForeground(textColor);
        txtRecipientEmail.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createLineBorder(primaryColor, 2),
            BorderFactory.createEmptyBorder(5, 10, 5, 10)
        ));
        
        btnBrowse = new JButton("Browse Attendance File");
        btnBrowse.setFont(new Font("Segoe UI", Font.BOLD, 16));
        btnBrowse.setBackground(secondaryColor);
        btnBrowse.setForeground(Color.WHITE);
        btnBrowse.setFocusPainted(false);
        btnBrowse.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createRaisedBevelBorder(),
            BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));
        btnBrowse.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnBrowse.setContentAreaFilled(true);
        btnBrowse.setOpaque(true);
        
        btnProcess = new JButton("Process Summary Email");
        btnProcess.setFont(new Font("Segoe UI", Font.BOLD, 16));
        btnProcess.setBackground(accentColor);
        btnProcess.setForeground(Color.WHITE);
        btnProcess.setFocusPainted(false);
        btnProcess.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createRaisedBevelBorder(),
            BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));
        btnProcess.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnProcess.setContentAreaFilled(true);
        btnProcess.setOpaque(true);
        
        btnSendEmail = new JButton("Send Email");
        btnSendEmail.setFont(new Font("Segoe UI", Font.BOLD, 16));
        btnSendEmail.setBackground(primaryColor);
        btnSendEmail.setForeground(Color.WHITE);
        btnSendEmail.setFocusPainted(false);
        btnSendEmail.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createRaisedBevelBorder(),
            BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));
        btnSendEmail.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnSendEmail.setContentAreaFilled(true);
        btnSendEmail.setOpaque(true);
        
        txtResult = new JTextArea(20, 70);
        txtResult.setFont(new Font("Consolas", Font.PLAIN, 14));
        txtResult.setEditable(false);
        txtResult.setBackground(Color.WHITE);
        txtResult.setForeground(textColor);
        scrollPane = new JScrollPane(txtResult);
        scrollPane.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createTitledBorder("Activity Log"),
            BorderFactory.createEmptyBorder(5, 5, 5, 5)
        ));

        previewPane = new JEditorPane();
        previewPane.setContentType("text/html");
        previewPane.setEditable(false);
        previewPane.setBackground(Color.WHITE);
        previewPane.setForeground(textColor);

        // Initialize table for email content
        String[] columnNames = { "SR.NO", "NAME", "DEPARTMENT", "DAY 1 STATUS", "DAY 2 STATUS" };
        Object[][] data = {};
        emailTable = new JTable(data, columnNames);
        emailTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        emailTable.setRowHeight(25);
        emailTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        emailTable.getTableHeader().setBackground(primaryColor);
        emailTable.getTableHeader().setForeground(Color.WHITE);
        emailTable.setSelectionBackground(secondaryColor);
        emailTable.setSelectionForeground(Color.WHITE);
        tableScrollPane = new JScrollPane(emailTable);
        tableScrollPane.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createTitledBorder("NOT OK Employees Table"),
            BorderFactory.createEmptyBorder(5, 5, 5, 5)
        ));

        // Create tabbed pane for result display
        resultTabbedPane = new JTabbedPane();
        resultTabbedPane.setFont(new Font("Segoe UI", Font.BOLD, 14));
        resultTabbedPane.addTab("Log", scrollPane);
        resultTabbedPane.addTab("Email Preview", new JScrollPane(previewPane));
        resultTabbedPane.addTab("NOT OK Table", tableScrollPane);

        // Make buttons initially disabled until file is selected
        btnProcess.setEnabled(false);
        btnSendEmail.setEnabled(false);
    }

    private void setupLayout() {
        setLayout(new BorderLayout());
        
        // Create header panel
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBackground(primaryColor);
        headerPanel.setBorder(BorderFactory.createEmptyBorder(15, 20, 15, 20));
        
        JLabel titleLabel = new JLabel("Attendance Tracking & Email Reporting System");
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 24));
        titleLabel.setForeground(Color.WHITE);
        headerPanel.add(titleLabel, BorderLayout.CENTER);
        
        add(headerPanel, BorderLayout.NORTH);

        // Main content panel
        JPanel mainPanel = new JPanel(new BorderLayout());
        mainPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));
        mainPanel.setBackground(backgroundColor);
        
        // Top panel for file selection
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(backgroundColor);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);

        gbc.gridx = 0;
        gbc.gridy = 0;
        JLabel lblAttendanceFile = new JLabel("Attendance File:");
        lblAttendanceFile.setFont(new Font("Segoe UI", Font.BOLD, 16));
        lblAttendanceFile.setForeground(textColor);
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
        middlePanel.setBackground(backgroundColor);
        gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);

        gbc.gridx = 0;
        gbc.gridy = 0;
        JLabel lblRecipientEmail = new JLabel("Recipient Email:");
        lblRecipientEmail.setFont(new Font("Segoe UI", Font.BOLD, 16));
        lblRecipientEmail.setForeground(textColor);
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

        // Add panels to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);
        mainPanel.add(middlePanel, BorderLayout.CENTER);
        
        // Add main panel to frame
        add(mainPanel, BorderLayout.CENTER);

        // Bottom panel for filters
        JPanel filterPanel = new JPanel(new GridBagLayout());
        filterPanel.setBackground(backgroundColor);
        filterPanel.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createTitledBorder("Filter Options"),
            BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));
        GridBagConstraints gbc2 = new GridBagConstraints();
        gbc2.insets = new Insets(5, 5, 5, 5);

        // Attendance Status Filter
        gbc2.gridx = 0;
        gbc2.gridy = 0;
        JLabel lblAttendanceFilter = new JLabel("Attendance Status Filter:");
        lblAttendanceFilter.setFont(new Font("Segoe UI", Font.BOLD, 14));
        lblAttendanceFilter.setForeground(textColor);
        filterPanel.add(lblAttendanceFilter, gbc2);

        String[] statusOptions = { "ALL", "PRESENT", "ABSENT" };
        this.cmbAttendanceStatus = new JComboBox<>(statusOptions);
        this.cmbAttendanceStatus.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        this.cmbAttendanceStatus.setBackground(Color.WHITE);
        this.cmbAttendanceStatus.setForeground(textColor);
        this.cmbAttendanceStatus.setPreferredSize(new Dimension(150, 35));
        gbc2.gridx = 1;
        filterPanel.add(this.cmbAttendanceStatus, gbc2);

        String[] dayOptions = { "Day 1", "Day 2" };
        this.cmbDayFilter = new JComboBox<>(dayOptions);
        this.cmbDayFilter.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        this.cmbDayFilter.setBackground(Color.WHITE);
        this.cmbDayFilter.setForeground(textColor);
        this.cmbDayFilter.setPreferredSize(new Dimension(120, 35));
        gbc2.gridx = 2;
        filterPanel.add(this.cmbDayFilter, gbc2);

        JButton btnApplyStatusFilter = new JButton("Apply Status Filter");
        btnApplyStatusFilter.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btnApplyStatusFilter.setBackground(secondaryColor);
        btnApplyStatusFilter.setForeground(Color.WHITE);
        btnApplyStatusFilter.setFocusPainted(false);
        btnApplyStatusFilter.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnApplyStatusFilter.setPreferredSize(new Dimension(170, 35));
        btnApplyStatusFilter.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createRaisedBevelBorder(),
            BorderFactory.createEmptyBorder(5, 10, 5, 10)
        ));
        btnApplyStatusFilter.setContentAreaFilled(true);
        btnApplyStatusFilter.setOpaque(true);
        gbc2.gridx = 3;
        filterPanel.add(btnApplyStatusFilter, gbc2);

        // Absenteeism Filter
        gbc2.gridx = 4;
        gbc2.gridy = 0;
        JLabel lblAbsenteeismFilter = new JLabel("Absenteeism Filter:");
        lblAbsenteeismFilter.setFont(new Font("Segoe UI", Font.BOLD, 14));
        lblAbsenteeismFilter.setForeground(textColor);
        filterPanel.add(lblAbsenteeismFilter, gbc2);

        String[] absenteeismOptions = { "ALL", "Absent on Day 1 Only", "Absent on Day 2 Only", "Absent on Both Days",
                "Absent on Any Day" };
        this.cmbAbsenteeismType = new JComboBox<>(absenteeismOptions);
        this.cmbAbsenteeismType.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        this.cmbAbsenteeismType.setBackground(Color.WHITE);
        this.cmbAbsenteeismType.setForeground(textColor);
        this.cmbAbsenteeismType.setPreferredSize(new Dimension(200, 35));
        gbc2.gridx = 5;
        filterPanel.add(this.cmbAbsenteeismType, gbc2);

        JButton btnApplyAbsenteeismFilter = new JButton("Apply Absenteeism Filter");
        btnApplyAbsenteeismFilter.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btnApplyAbsenteeismFilter.setBackground(accentColor);
        btnApplyAbsenteeismFilter.setForeground(Color.WHITE);
        btnApplyAbsenteeismFilter.setFocusPainted(false);
        btnApplyAbsenteeismFilter.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnApplyAbsenteeismFilter.setPreferredSize(new Dimension(200, 35));
        btnApplyAbsenteeismFilter.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createRaisedBevelBorder(),
            BorderFactory.createEmptyBorder(5, 10, 5, 10)
        ));
        btnApplyAbsenteeismFilter.setContentAreaFilled(true);
        btnApplyAbsenteeismFilter.setOpaque(true);
        gbc2.gridx = 6;
        filterPanel.add(btnApplyAbsenteeismFilter, gbc2);

        // Add action listeners for filter buttons
        btnApplyStatusFilter.addActionListener(e -> applyAttendanceStatusFilter(
                this.cmbAttendanceStatus.getSelectedItem().toString(),
                this.cmbDayFilter.getSelectedItem().toString()));

        btnApplyAbsenteeismFilter.addActionListener(e -> applyAbsenteeismFilter(
                this.cmbAbsenteeismType.getSelectedItem().toString()));

        add(filterPanel, BorderLayout.SOUTH);
        
        // Create a separate panel for the result area to avoid layout conflicts
        JPanel resultPanel = new JPanel(new BorderLayout());
        resultPanel.setBackground(backgroundColor);
        resultPanel.setBorder(BorderFactory.createEmptyBorder(10, 20, 20, 20));
        resultPanel.add(resultTabbedPane, BorderLayout.CENTER);
        
        // Add result panel to the main frame
        add(resultPanel, BorderLayout.SOUTH);
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
            JOptionPane.showMessageDialog(this, "Please select an attendance file first.", "Error",
                    JOptionPane.ERROR_MESSAGE);
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
            txtResult.append("\nEmail Content Preview updated in 'Email Preview' tab.\n");
            // txtResult.append(emailContent + "\n");

            // Set HTML preview
            previewPane.setText(emailContent);

            // Switch to preview tab
            resultTabbedPane.setSelectedIndex(1);

            // Populate legacy table
            populateEmailTableFromEmployees(employees);

            // Update master sheet with processed data
            try {
                controller.updateMasterSheet();
                txtResult.append("Master sheet updated successfully with employee data\n");
            } catch (Exception ex) {
                txtResult.append("Error updating master sheet: " + ex.getMessage() + "\n");
            }

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
            JOptionPane.showMessageDialog(this, "Please enter recipient email address.", "Error",
                    JOptionPane.ERROR_MESSAGE);
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
            JOptionPane.showMessageDialog(this, "Please enter recipient email address.", "Error",
                    JOptionPane.ERROR_MESSAGE);
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

            // Note: Master sheet was already updated after processing attendance data

        } else {
            txtResult.append("Failed to send email.\n");
            JOptionPane.showMessageDialog(this, "Failed to send email. Please check your email settings.",
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void applyAttendanceStatusFilter(String status, String day) {
        if (employees == null || employees.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please process an attendance file first.", "Error",
                    JOptionPane.ERROR_MESSAGE);
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
            java.util.List<model.Employee> filteredEmployees = controller.filterEmployeesByStatus(attendanceStatus,
                    dayValue);

            txtResult.append("\nFilter applied: " + status + " employees on " + day + "\n");
            txtResult.append("Number of employees: " + filteredEmployees.size() + "\n");

            if (!filteredEmployees.isEmpty()) {
                txtResult.append("Employee Names: " + String.join(", ",
                        filteredEmployees.stream().map(model.Employee::getName).toArray(String[]::new)) + "\n");
            }

        } catch (Exception ex) {
            txtResult.append("Error applying attendance status filter: " + ex.getMessage() + "\n");
            JOptionPane.showMessageDialog(this, "Error applying filter: " + ex.getMessage(), "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void applyAbsenteeismFilter(String absenteeismType) {
        if (employees == null || employees.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please process an attendance file first.", "Error",
                    JOptionPane.ERROR_MESSAGE);
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
            JOptionPane.showMessageDialog(this, "Error applying filter: " + ex.getMessage(), "Error",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void populateEmailTableFromEmployees(java.util.List<model.Employee> employees) {
        if (employees == null)
            return;

        java.util.List<Object[]> tableData = new java.util.ArrayList<>();

        for (model.Employee emp : employees) {
            String day1Attendance = emp.getDay1Attendance();
            String day2Attendance = emp.getDay2Attendance();
            
            // Check if employee is NOT OK (absent on one or both days)
            boolean isDay1Present = day1Attendance != null && 
                (day1Attendance.toUpperCase().contains("P") || day1Attendance.toUpperCase().contains("PRESENT"));
            boolean isDay2Present = day2Attendance != null && 
                (day2Attendance.toUpperCase().contains("P") || day2Attendance.toUpperCase().contains("PRESENT"));
            
            if (!(isDay1Present && isDay2Present)) {  // If not present on both days, then NOT OK
                String srNo = String.valueOf(emp.getSrNo());
                String name = emp.getName();
                String dept = emp.getTrade() != null ? emp.getTrade() : "N/A";
                String day1Status = day1Attendance != null ? day1Attendance : "N/A";
                String day2Status = day2Attendance != null ? day2Attendance : "N/A";

                tableData.add(new Object[] { srNo, name, dept, day1Status, day2Status });
            }
        }

        // Update the table with the parsed data
        Object[][] data = tableData.toArray(new Object[tableData.size()][]);
        String[] columnNames = { "SR.NO", "NAME", "DEPARTMENT", "DAY 1 STATUS", "DAY 2 STATUS" };

        // Create a new table model with the data
        javax.swing.table.DefaultTableModel model = new javax.swing.table.DefaultTableModel(data, columnNames) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // Make table non-editable
            }
        };

        emailTable.setModel(model);
        
        // Adjust column widths
        if (emailTable.getColumnModel().getColumnCount() > 0) {
            emailTable.getColumnModel().getColumn(0).setPreferredWidth(40);  // SR.NO
            emailTable.getColumnModel().getColumn(1).setPreferredWidth(150); // NAME
            emailTable.getColumnModel().getColumn(2).setPreferredWidth(120); // DEPARTMENT
            emailTable.getColumnModel().getColumn(3).setPreferredWidth(80);  // DAY 1 STATUS
            emailTable.getColumnModel().getColumn(4).setPreferredWidth(80);  // DAY 2 STATUS
        }
    }
}