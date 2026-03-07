/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt 
to change this license
 */

// This tells which folder structure the class belongs to
package com.mycompany.motorph_ms2_grp9; 

// ================================
// Java Standard Libraries
// ================================
// File handling
import java.io.FileInputStream; // Allows the program to Open and Read files 

// Date and formatting
import java.util.Date; // Allows to store date values, (used for birthday)
    // Allows to format date to readable format (deafault format - Wed Oct 21 00:00:00 PST 1987)
import java.text.SimpleDateFormat; 

// User input
import java.util.Scanner; // Allows userinput from keyboard

// =====================================
// Java Time API (Work Hour Calculation)
// =====================================
import java.time.LocalTime; //Stores time values (e.g., time in/out)
import java.time.Duration; // Calculates difference between two time

// ====================================================
// Third-Party Libraries (Apache POI - Excel Handling)
// ====================================================
import org.apache.poi.ss.usermodel.*; // Access Excel components (Workbook, Sheet, Row, Cell)
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // Specifically handles .xlsx files.

/**
 * 
 *
 * @author Kim
 */
public class MotorPH_MS2_Grp9 {
    
    // ======================================================
    //              *** DEFINED METHODS ***
    // ======================================================
    
    // ======================================================
    // EXCEL FILE HANDLING METHODS
    // ======================================================

    // Opens the Excel workbook file
    public static Workbook loadWorkbook(String filePath) {

        try {
            FileInputStream file = new FileInputStream(filePath);
            return new XSSFWorkbook(file);

        } catch (Exception e) {
            System.out.println("Error loading Excel file.");
            e.printStackTrace();
            return null;
            }
    }
    
    // Selects which sheet from the workbook will be used
    public static Sheet getSheet(Workbook workbook, int index) {

        if (workbook == null) return null;
            return workbook.getSheetAt(index);
    }
    
    // ======================================================
    // SALARY COMPUTATION METHODS
    // ======================================================
    
    // Calculates Gross Salary
    public static double processGrossSalary(double hourlyRate, double totalHoursworked){
        double grossSalary = hourlyRate * totalHoursworked;
        return grossSalary;
    }
    
    // ======================================================
    // GOVERNMENT CONTRIBUTION CALCULATIONS
    // ======================================================
    
    // SSS Contribution
    public static double calculateSss(double totalGross) {

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
        
        if (totalGross < start) {
            return minContribution;
            
        } else if (totalGross >= end) {
            return maxContribution;
            
        } else if (totalGross >= start && totalGross < end) {
            
            do {
                
                counter++;
                
                startRange = startRange + compRangeAdder;
                endRange = endRange + compRangeAdder;
                finalContribution = finalContribution + contAdder;
                
            } while (totalGross > endRange);
        
            return finalContribution;
            
         }  else {
            return 0;
        }
    }

    // PhilHealth Contribution
    public static double calculatePhilhealth(double totalGross) {

        double minSalary = 10000;
        double maxSalary = 60000;
        double rate = 0.03;

        double salaryBase;

        if (totalGross < minSalary) {
            salaryBase = minSalary;
            
        } else if (totalGross > maxSalary) {
            salaryBase = maxSalary;
            
        } else {
            salaryBase = totalGross;
        }

        double totalPremium = salaryBase * rate;
        return totalPremium / 2;
    }

    // Pag-IBIG Contribution
     public static double calculatePagibig(double totalGross) {

        double employeeRate;

        if (totalGross >= 1000 && totalGross <= 1500) {
            employeeRate = 0.01;   // 1%
            
        } else if (totalGross > 1500) {
            employeeRate = 0.02;   // 2%
            
        } else {
            return 0; // below 1000, no contribution
        }

        double pagIbig = totalGross * employeeRate;

        // Apply maximum cap of 100
        if (pagIbig > 100) {
            pagIbig = 100;
        }

        return pagIbig;
    }
        
    // ======================================================
    // TAX CALCULATION
    // ======================================================

    // Computes withholding tax based on taxable income
    public static double calculateWithHoldingTax(double taxableIncome) {

        if (taxableIncome <= 20832) {
            return 0;
        } 
        else if (taxableIncome < 33333) {
            return (taxableIncome - 20833) * 0.20;
        } 
        else if (taxableIncome < 66667) {
            return 2500 + (taxableIncome - 33333) * 0.25;
        } 
        else if (taxableIncome < 166667) {
            return 10833 + (taxableIncome - 66667) * 0.30;
        } 
        else if (taxableIncome < 666667) {
            return 40833.33 + (taxableIncome - 166667) * 0.32;
        } 
        else {
            return 200833.33 + (taxableIncome - 666667) * 0.35;
        }
    }
    
    // ======================================================
    // PAYROLL OUTPUT DISPLAY
    // ======================================================

    // Prints formatted payroll breakdown
    public static void printFormattedPayroll(
            int id,
            String employeeName,
            String birthday,
            int month,
            double firstHours,
            double secondHours,
            double firstGross,
            double secondGross,
            double sssContribution,
            double philhealthContribution,
            double pagibigContribution,
            double totalDeductions,
            double withHoldingTax,
            double netSalary) {

        String monthName = java.time.Month.of(month).name();

        System.out.println("\n======================================");
        System.out.println("Employee Number: " + id);
        System.out.println("Employee Name: " + employeeName);
        System.out.println("Birthday: " + birthday);

        System.out.println("\nCutoff Date: " + monthName + 
                " 1 to " + monthName + " 15");
        System.out.println("Total Hours Worked: " + firstHours);
        System.out.println("Gross Salary: " + firstGross);
        System.out.println("Net Salary: " + firstGross);

        System.out.println("\nCutoff Date: " + monthName +
                " 16 to " + monthName +
                " " + java.time.YearMonth.of(2024, month).lengthOfMonth());

        System.out.println("Total Hours Worked: " + secondHours);
        System.out.println("Gross Salary: " + secondGross);
        System.out.println("Each Deduction");
        System.out.println("SSS: " + sssContribution);
        System.out.println("PhilHealth: " + philhealthContribution);
        System.out.println("Pag-IBIG: " + pagibigContribution);
        System.out.println("Tax: " + withHoldingTax);
        System.out.println("Total Deductions: " + totalDeductions);
        System.out.println("Net Salary: " + netSalary);
        System.out.println("======================================");
    }

    // ======================================================
    // MAIN PAYROLL GENERATION LOGIC
    // ======================================================

    public static void generatePayrollForEmployee(
            int employeeId,
            String employeeName,
            String formattedBirthday,
            double hourlyRate,
            Sheet attendanceSheet) {

    // COMPUTE PAYROLL
        for (int month = 6; month <= 12; month++) {

            Duration firstCutoff = Duration.ZERO;
            Duration secondCutoff = Duration.ZERO;

            for (Row attendanceRow : attendanceSheet) {

                if (attendanceRow.getRowNum() == 0) 
                    continue; //skip column headers, to avoid reading text values
                
                //Gets employee ID from attendance
                int attendanceEmpId =
                        (int) attendanceRow.getCell(0).getNumericCellValue();
                
                //Ensures to process only the selected employee's attendance
                if (attendanceEmpId != employeeId) 
                    continue;
                
                //Gets the attendance date
                java.time.LocalDate attendanceDate =
                        attendanceRow.getCell(3)
                                .getLocalDateTimeCellValue()
                                .toLocalDate();
                
                if (attendanceDate.getMonthValue() != month)
                    continue;

                //Reads login time
                LocalTime login =
                        attendanceRow.getCell(4)
                                .getLocalDateTimeCellValue()
                                .toLocalTime();
                
                //Reads login time
                LocalTime logout =
                        attendanceRow.getCell(5)
                                .getLocalDateTimeCellValue()
                                .toLocalTime();
                
                //Define Official Work Schedule
                LocalTime officialStart = LocalTime.of(8, 0);
                LocalTime graceLimit = LocalTime.of(8, 5);
                LocalTime officialEnd = LocalTime.of(17, 0);
                
                //Apply Grace Period Rule
                if (login.isBefore(officialStart) || !login.isAfter(graceLimit))
                    login = officialStart;
                
                //Limit Logout Time
                if (logout.isAfter(officialEnd))
                    logout = officialEnd;
                
                //Skip Invalid Time Records
                if (logout.isBefore(officialStart) || login.isAfter(officialEnd))
                    continue;
                
                //Calculate Daily Work Duration
                Duration daily = Duration.between(login, logout);
                
                //Deduct Lunch Break
                if (daily.compareTo(Duration.ofHours(1)) > 0)
                    daily = daily.minusHours(1);
                
                else
                    daily = Duration.ZERO;//If work time is less than 1 hour → set to 0.
                
                //Assign Hours to Payroll Cutoff
                if (attendanceDate.getDayOfMonth() <= 15)
                    firstCutoff = firstCutoff.plus(daily);
                
                else
                    secondCutoff = secondCutoff.plus(daily);
            }

            //Convert Duration to Hours
            double firstHours = firstCutoff.toMinutes() / 60.0;
            double secondHours = secondCutoff.toMinutes() / 60.0;
            
            //Skip Months Without Attendance
            if (firstHours == 0 && secondHours == 0)
                continue;

            double firstGross = processGrossSalary(hourlyRate, firstHours);
            double secondGross = processGrossSalary(hourlyRate, secondHours);

            double totalGross = firstGross + secondGross;

            double sss = calculateSss(totalGross);
            double philhealth = calculatePhilhealth(totalGross);
            double pagibig = calculatePagibig(totalGross);
            
            // Government deductions
            double govDeductions = sss + philhealth + pagibig;

            // Compute taxable income
            double taxableIncome = totalGross - govDeductions;

            // Compute withholding tax
            double tax = calculateWithHoldingTax(taxableIncome);

            // Final deductions including tax
            double totalDeductions = govDeductions + tax;

            // Net salary deducted from second cutoff
            double netSalary = secondGross - totalDeductions;

            printFormattedPayroll( //Display Payroll
                    employeeId,
                    employeeName,
                    formattedBirthday,
                    month,
                    firstHours,
                    secondHours,
                    firstGross,
                    secondGross,
                    sss,
                    philhealth,
                    pagibig,
                    totalDeductions,
                    tax,
                    netSalary
            );
        }
    }
        
    public static void main(String[] args) { // Program Entry point
        
        // ======================================================
        // LOGIN SECTION
        // ======================================================
        
        // Calling the import scanner, this will allow program to accept input
        Scanner scanner = new Scanner(System.in); 
        
        System.out.println("==================================");
        System.out.println("       WELCOME TO MOTORPH"         );
        System.out.println("==================================");
        System.out.print("Username: "); //Prompts the user to enter something
            String inputUsername = scanner.nextLine();

        System.out.print("Password: "); //Prompts the user to enter something
            String inputPassword = scanner.nextLine();
            
        
        System.out.println("\n=== Log In as Employee ===");
        System.out.println("Username: " + inputUsername);
        System.out.println("Password: " + inputPassword);
          
        // ======================================================
        // EMPLOYEE LOGIN
        // ======================================================
           
            if (inputUsername.equals("employee") && inputPassword.equals("12345")) {
            
            
            System.out.println("\nLogin Successful!");
                
            System.out.print("Enter Employee ID: ");

            while (!scanner.hasNextInt()) {
                System.out.println("Invalid input. Please enter a numeric Employee ID.");
                scanner.next(); // Invalid input
            }

            int inputId = scanner.nextInt();

            boolean found = false;

            try {

                Workbook workbook = loadWorkbook("CP1_Grp9_MotorPH_Employee Data.xlsx");
                if (workbook == null) return;  // stop program if file failed

                Sheet employeeSheet = getSheet(workbook, 0);
                Sheet attendanceSheet = getSheet(workbook, 1);

                DataFormatter formatter = new DataFormatter(); // To modify formats
                boolean firstRow = true; // skip header properly

                // =========================
                // SEARCH EMPLOYEE RECORD
                // =========================
                
                for (Row row : employeeSheet) {

                    if (firstRow) {
                        firstRow = false;
                        continue;
                    }

                    int id = Integer.parseInt(
                            formatter.formatCellValue(row.getCell(0)));

                    if (id == inputId) {

                        String lastName = formatter.formatCellValue(row.getCell(1));
                        String firstName = formatter.formatCellValue(row.getCell(2));
                        String employeeName = firstName + " " + lastName;

                        Cell birthdayCell = row.getCell(3);
                        Date birthday = birthdayCell.getDateCellValue();

                        SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                        String formattedBirthday = sdf.format(birthday);

                        System.out.println("\n=== MotorPH Employee Details ===");
                        System.out.println("Employee Number: " + id);
                        System.out.println("Employee Name: " + employeeName);
                        System.out.println("Birthday: " + formattedBirthday);

                        found = true;
                        break;            
                    }
                }

                if (!found) {
                    System.out.println("Employee number does not exist");
                }

                workbook.close();

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
  
        // ======================================================
        // PAYROLL STAFF LOGIN
        // ======================================================

        else if (inputUsername.equals("payroll_staff") 
                && inputPassword.equals("12345")) {

            System.out.println("\nLogin Successful!");

            try {

                Workbook workbook = 
                        loadWorkbook("CP1_Grp9_MotorPH_Employee Data.xlsx");
            
                if (workbook == null) return;

                Sheet employeeSheet = workbook.getSheetAt(0);
                Sheet attendanceSheet = workbook.getSheetAt(1);

                // =========================
                // MAIN MENU
                // =========================
                
                System.out.println("\n=== MotorPH Payroll System ===");
                System.out.println("1. Process Payroll");
                System.out.println("2. Exit");
                System.out.print("Select option: ");
                    
                    // VALIDATE MAIN MENU INPUT
                    while (!scanner.hasNextInt()) {
                        System.out.println("Invalid input. Please enter the option number.");
                        scanner.next();
                    }

                    int inputOption1 = scanner.nextInt();
                    scanner.nextLine();

                // ======================================================
                // PROCESS PAYROLL
                // ======================================================
                
                if (inputOption1 == 1) {

                    System.out.println("\n=== Process Payroll ===");
                    System.out.println("1. One Employee");
                    System.out.println("2. All Employees");
                    System.out.println("3. Exit");
                    System.out.print("Select option: ");
                    
                int inputOption2;

                // VALIDATE SECOND MENU INPUT
                while (!scanner.hasNextInt()) {
                    System.out.println("Invalid input. Please enter the option number.");
                    scanner.next();
                }

                inputOption2 = scanner.nextInt();
                scanner.nextLine();

                    // =========================================
                    // ONE EMPLOYEE PAYROLL
                    // =========================================
                    
                    if (inputOption2 == 1) {

                        System.out.print("\nEnter Employee ID: ");
                       
                       // VALIDATE EMPLOYEE ID INPUT
                       while (!scanner.hasNextInt()) {
                       System.out.println("Invalid input. Please enter a numeric Employee ID.");
                       scanner.next();
                       }

                         int inputId = scanner.nextInt();

                        boolean found = false;

                        for (Row row : employeeSheet) {

                            if (row.getRowNum() == 0) continue;

                                int id = (int) row.getCell(0)
                                        .getNumericCellValue();

                            if (id == inputId) {

                                String lastName = 
                                        row.getCell(1).getStringCellValue();
                                String firstName = 
                                        row.getCell(2).getStringCellValue();
                                String employeeName = 
                                        firstName + " " + lastName;

                                Date birthday = 
                                        row.getCell(3).getDateCellValue();
                                SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                                String formattedBirthday = sdf.format(birthday);

                                double hourlyRate = row.getCell(18).getNumericCellValue();

                                generatePayrollForEmployee(
                                        id,
                                        employeeName,
                                        formattedBirthday,
                                        hourlyRate,
                                        attendanceSheet
                                );

                                found = true;
                                break;
                            }
                        }

                        if (!found) {
                            System.out.println("Employee ID Not Found.");
                        }
                    }

                    // =========================================
                    // ALL EMPLOYEES PAYROLL
                    // =========================================
                    
                    else if (inputOption2 == 2) {

                        for (Row row : employeeSheet) {

                            if (row.getRowNum() == 0) continue;

                            int id = (int) row.getCell(0).getNumericCellValue();

                            String lastName = row.getCell(1).getStringCellValue();
                            String firstName = row.getCell(2).getStringCellValue();
                            String employeeName = firstName + " " + lastName;

                            Date birthday = row.getCell(3).getDateCellValue();
                            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                            String formattedBirthday = sdf.format(birthday);

                            double hourlyRate = 
                                    row.getCell(18).getNumericCellValue();

                            generatePayrollForEmployee(
                                    id,
                                    employeeName,
                                    formattedBirthday,
                                    hourlyRate,
                                    attendanceSheet
                            );
                        }
                    }

                    else if (inputOption2 == 3) {
                        System.out.println("Exiting Process Payroll...");
                    }

                    else {
                        System.out.println("Invalid option.");
                    }

                }

                else if (inputOption1 == 2) {
                        System.out.println("Exiting program...");
                }

            else {
                System.out.println("Invalid option.");
                }

                    workbook.close();

                } 
                catch (Exception e) {
                    e.printStackTrace();
                }
                
        }
        else {
            System.out.println("Invalid username or password.");
        }
      }
    }

