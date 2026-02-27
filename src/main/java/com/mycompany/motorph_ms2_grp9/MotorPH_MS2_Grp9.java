/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.motorph_ms2_grp9; // This tells which folder structure the class belongs to

import java.io.FileInputStream; // Allows the program to Open and Read files 
import org.apache.poi.ss.usermodel.*; // * Means import everything inside this package
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // Specifically handles .xlsx files.
import java.util.Date; // Allows to store date values, in this program it was used for birthday
import java.text.SimpleDateFormat; // Allows to format date into readable format - deafault format - Wed Oct 21 00:00:00 PST 1987
import java.util.Scanner; // Allows userinput from keyboard

/**
 *
 * @author Kim
 */
public class MotorPH_MS2_Grp9 {
    
    public static double calculateTotalhours(double logOut, double logIn){
        double totalHoursworked = logOut - logIn;
        return totalHoursworked;
    }
    
    public static double processGrosssalary(double hourlyRate, double totalHoursworked){
        double grossSalary = hourlyRate * totalHoursworked;
        return grossSalary;
    }
    
    public static double processNetsalary(double grossSalary, double totalDeductions){
        double netSalary = grossSalary - totalDeductions;
        return netSalary;
    }
    
    public static double calculateSss_ForLoop(double grossSalary) {

        int start = 3250;
        int end = 24750;
        double minContribution = 135.00;
        double maxContribution = 1125.00;
        
        //Used in do while loop
        int counter = 0;
        int startRange = 3250;
        int endRange = 3750;
        int compRangeAdder = 500;
        double finalContribution = 157.50;
        double contAdder = 22.50;
        
        if (grossSalary < start) {
            return minContribution;
        } else if (grossSalary >= end) {
            return maxContribution;
        } else if (grossSalary >= start && grossSalary < end) {
            do {
                System.out.println("Your counter is: " + counter);
                System.out.println("Your contribution is: " + finalContribution);
                System.out.println("Your startRange is: " + startRange);
                System.out.println("Your endRange is: " + endRange);
                System.out.println("Your grossSalary is: " + grossSalary + "\n");
                
                counter++;
                
                startRange = startRange + compRangeAdder;
                endRange = endRange + compRangeAdder;
                finalContribution = finalContribution + contAdder;
                
            } while (grossSalary > endRange);
        
            return finalContribution;
         }  else {
            return 0;
        }
    }
        
//    public static double calculatePhilhealth(double grossSalary) {
//
//        } 
//
//    public static double calculatePagibig(double grossSalary) {
//
//        }  
//    
    
//    public static double calculateTax(double grossSalary) {
//
//        }  
//        
        
    public static void main(String[] args) { // Program Entry point
        
        //Log In
        Scanner scanner = new Scanner(System.in); // Calling the import scanner, this will allow program to accept input

        System.out.print("Username: "); //Prompts the user to enter something
        String inputUsername = scanner.nextLine();

        System.out.print("Password: "); //Prompts the user to enter something
        String inputPassword = scanner.nextLine();
        
                System.out.println("\n--- Log In to MotorPH ---");
                System.out.println("Username: " + inputUsername);
                System.out.println("Password: " + inputPassword);
          
    //If username is employee            
        if (inputUsername.equals("employee") && inputPassword.equals("12345")) {

            System.out.println("\nLogin Successful!");
                
        System.out.print("Enter Employee ID: "); //Prompts the user to enter something
        int inputId = scanner.nextInt();

        boolean found = false;

    try {

        FileInputStream dataBase = new FileInputStream("CP1_Grp9_MotorPH_Employee Data.xlsx"); //Location of the file reference
        Workbook employeeDetails = new XSSFWorkbook(dataBase); //Enabling workbooks in xlxs file
        Sheet sheet = employeeDetails.getSheetAt(0); // Allows to select which specific worksheet to cath details

        DataFormatter formatter = new DataFormatter(); // To modify formats
        boolean firstRow = true; // skip header properly

        for (Row row : sheet) {

            if (firstRow) {
                firstRow = false;
                continue;
            }

            int id = Integer.parseInt(formatter.formatCellValue(row.getCell(0)));

            if (id == inputId) {

                String lastName = formatter.formatCellValue(row.getCell(1));
                String firstName = formatter.formatCellValue(row.getCell(2));
                String employeeName = firstName + " " + lastName;

                Cell birthdayCell = row.getCell(3);
                Date birthday = birthdayCell.getDateCellValue();

                SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                String formattedBirthday = sdf.format(birthday);

                System.out.println("\n--- MotorPH Employee Details ---");
                System.out.println("Employee Number: " + id);
                System.out.println("Employee Name: " + employeeName);
                System.out.println("Birthday: " + formattedBirthday);
                System.out.println("--------------------------------");

                found = true;
                break;            }
        }

        if (!found) {
            System.out.println("Employee ID Not found.");
        }

        employeeDetails.close();
        dataBase.close();

    } 
    catch (Exception e) {
            e.printStackTrace();
        }
  
    //If username is payroll_staff
} else if (inputUsername.equals("payroll_staff") && inputPassword.equals("12345")) {

    System.out.println("\nLogin Successful!");
    
    System.out.print("Process Payroll: "); //Prompts the user to enter something
        String inputProcess = scanner.nextLine();
        
    System.out.print("Enter Employee ID: "); //Prompts the user to enter something
        int inputId = scanner.nextInt();
        
         boolean found = false;
    
    try {

        FileInputStream dataBase = new FileInputStream("CP1_Grp9_MotorPH_Employee Data.xlsx"); //Location of the file reference
        Workbook employeeDetails = new XSSFWorkbook(dataBase); //Enabling workbooks in xlxs file
        Sheet sheet = employeeDetails.getSheetAt(0); // Allows to select which specific worksheet to cath details

        DataFormatter formatter = new DataFormatter(); // To modify formats
        boolean firstRow = true; // skip header properly

        for (Row row : sheet) {

            if (firstRow) {
                firstRow = false;
                continue;
            }

            int id = Integer.parseInt(formatter.formatCellValue(row.getCell(0)));

            if (id == inputId) {

                String lastName = formatter.formatCellValue(row.getCell(1));
                String firstName = formatter.formatCellValue(row.getCell(2));
                String employeeName = firstName + " " + lastName;

                Cell birthdayCell = row.getCell(3);
                Date birthday = birthdayCell.getDateCellValue();

                SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                String formattedBirthday = sdf.format(birthday);
//1st Cut off


                System.out.println("\n--- MotorPH Payroll ---");
                System.out.println("Employee Number: " + id);
                System.out.println("Employee Name: " + employeeName);
                System.out.println("Birthday: " + formattedBirthday);
                System.out.println("Cutoff Date: June 1 to June 15");
                System.out.println("Total Hours Worked: " + formattedBirthday);
                System.out.println("Gross Salary: " + formattedBirthday);
                System.out.println("Net Salary: " + formattedBirthday);
                System.out.println("Cutoff Date: June 16 to June 30");
                System.out.println("Total Hours Worked: " + formattedBirthday);
                System.out.println("Gross Salary: " + formattedBirthday);
                System.out.println("Each Deduction: ");
                System.out.println("SSS: " + formattedBirthday);
                System.out.println("PhilHealth: " + formattedBirthday);
                System.out.println("Pag-IBIG: " + formattedBirthday);
                System.out.println("Tax: " + formattedBirthday);
                System.out.println("Total Deductions: " + formattedBirthday);
                System.out.println("Net Salary: " + formattedBirthday);
                System.out.println("--------------------------------");

                found = true;
                break;            }
        }

        if (!found) {
            System.out.println("Employee ID Not found.");
        }

        employeeDetails.close();
        dataBase.close();

    } 
    catch (Exception e) {
            e.printStackTrace();
        }
    
    
    
} else {

    System.out.println("Invalid username or password.");
}
    }
}