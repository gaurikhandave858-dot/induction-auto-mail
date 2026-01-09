import java.util.ArrayList;
import java.util.List;

import model.EmailRecord;
import model.ExcelProcessor;

public class InitializeMasterSheet {        
    public static void main(String[] args) {
        try {
            ExcelProcessor excelProcessor = new ExcelProcessor();                       
            // Create an empty list to initialize the master sheet                                  
            List<EmailRecord> records = new ArrayList<>();                              
            // This will create the master sheet if it doesn't exist                                
            excelProcessor.updateMasterSheet(records, "master_attendance_summary.xlsx");
            System.out.println("Master attendance summary Excel sheet created successfully!");                                                  
            System.out.println("The sheet is located at: master_attendance_summary.xlsx");                                                      
            System.out.println("It will store all summary emails with both historical and recent data.");                                   
        } catch (Exception e) {
            System.err.println("Error creating master sheet: " + e.getMessage());                   
            e.printStackTrace();
        }
    }
}