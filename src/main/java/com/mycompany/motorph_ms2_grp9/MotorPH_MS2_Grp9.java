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
import java.text.SimpleDateFormat; // Allows to format date to readable format (deafault format - Wed Oct 21 00:00:00 PST 1987)

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
 * MotorPH Payroll Processing System
 *
 * This program reads employee and attendance records from an Excel file
 * and computes payroll information including:
 * - Gross salary
 * - Government deductions (SSS, PhilHealth, Pag-IBIG)
 * - Withholding tax
 * - Net salary
 *
 * The system supports two login roles:
 * 1. Employee – view employee details
 * 2. Payroll Staff – process payroll for one or all employees
 *
 * Apache POI is used for reading Excel files.
 *
 * Author: Kim
 */

public class MotorPH_MS2_Grp9 {
    
    // ======================================================
    //              *** DEFINED METHODS ***
    // ======================================================
    
    // ======================================================
    // CONSTANTS
    // ======================================================
    private static final LocalTime OFFICIAL_START = LocalTime.of(8, 0);
    private static final LocalTime GRACE_LIMIT = LocalTime.of(8, 5);
    private static final LocalTime OFFICIAL_END = LocalTime.of(17, 0);
    
    // ======================================================
    // INPUT VALIDATION
    // ======================================================
    public static int getValidatedIntInput(Scanner scanner, String message) {
        System.out.print(message);
        while (!scanner.hasNextInt()) {
            System.out.println("Invalid input. Please enter a number.");
            scanner.next();
        }
        return scanner.nextInt();
    }
    
    // ======================================================
    // EXCEL FILE HANDLING METHODS
    // ======================================================

    /**
    * Loads an Excel workbook from the specified file path.
    *
    * This method opens the Excel file using Apache POI and returns
    * a Workbook object that allows the program to access sheets,
    * rows, and cells within the Excel document.
    *
    * @param filePath the file location of the Excel workbook
    * @return Workbook object if the file is successfully loaded,
    *         otherwise returns null if an error occurs
    */
    public static Workbook loadWorkbook(String filePath) {

        try {
            FileInputStream file = new FileInputStream(filePath);
            return new XSSFWorkbook(file);

        } catch (Exception e) {
            System.out.println("Error loading Excel file.");
            return null;
            }
    }
    
    // Selects which sheet from the workbook will be used
    public static Sheet getSheet(Workbook workbook, int index) {
        if (workbook == null) return null;
            return workbook.getSheetAt(index);
    }
    
    public static java.util.Map<Integer, java.util.List<Row>> groupAttendanceByEmployee(Sheet attendanceSheet) {

    java.util.Map<Integer, java.util.List<Row>> groupedData = new java.util.HashMap<>();

    for (Row row : attendanceSheet) {

        if (row.getRowNum() == 0) continue;

        int empId = (int) row.getCell(0).getNumericCellValue();

        groupedData.putIfAbsent(empId, new java.util.ArrayList<>());
        groupedData.get(empId).add(row);
    }

    return groupedData;
}
    // ======================================================
    // CALCULATIONS
    // ======================================================
    
    // Calculates Gross Salary
    public static double calculateGrossSalary(double hourlyRate, double totalHoursWorked){
        return hourlyRate * totalHoursWorked;
    }
    
    // ======================================================
    // GOVERNMENT CONTRIBUTION CALCULATIONS
    // ======================================================
    
    /**
     * Calculates the employee's SSS (Social Security System) contribution
     * based on the total gross salary.
     *
     * The computation follows the SSS contribution table where salary
     * ranges determine the contribution amount. The contribution increases
     * incrementally depending on the salary bracket.
     *
     * @param totalGross the employee's total gross salary
     * @return the calculated SSS contribution amount
     */
    public static double calculateSss(double totalGross) {

        int start = 3250;
        int end = 24750;
        double minContribution = 135.00;
        double maxContribution = 1125.00;
        
        //Used in do while loop
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
                
                startRange = startRange + compRangeAdder;
                endRange = endRange + compRangeAdder;
                finalContribution = finalContribution + contAdder;
                
            } while (totalGross > endRange);
        
            return finalContribution;
            
         }  else {
            return 0;
        }
    }

    /**
    * Calculates the employee's PhilHealth contribution.
    *
    * The contribution is based on a percentage of the employee's salary.
    * A minimum and maximum salary base is applied to ensure that the
    * contribution falls within the allowable limits defined by PhilHealth.
    *
    * The computed premium is split equally between the employer
    * and employee, so this method returns only the employee share.
    *
    * @param totalGross the employee's total gross salary
    * @return the employee's PhilHealth contribution
    */
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

    /**
    * Calculates the employee's Pag-IBIG contribution.
    *
    * The contribution rate depends on the employee's salary range.
    * Employees earning between 1,000 and 1,500 contribute 1%,
    * while those earning above 1,500 contribute 2%.
    *
    * The contribution is capped at a maximum amount of 100 pesos.
    *
    * @param totalGross the employee's total gross salary
    * @return the employee's Pag-IBIG contribution
    */
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

    /**
    * Computes the withholding tax based on the employee's taxable income.
    *
    * The tax calculation follows the Philippine withholding tax brackets.
    * Different tax rates and base amounts apply depending on the
    * employee's taxable income range.
    *
    * @param taxableIncome the employee's income after government deductions
    * @return the computed withholding tax
    */
    public static double calculateWithholdingTax(double taxableIncome) {

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

    /**
    * Displays a formatted payroll report for an employee.
    *
    * The report includes employee information, hours worked per cutoff,
    * gross salary, government deductions, withholding tax,
    * total deductions, and final net salary.
    *
    * This method organizes payroll data into a readable format
    * for easier interpretation by payroll staff or employees.
    *
    * @param id employee identification number
    * @param employeeName employee full name
    * @param birthday employee birth date
    * @param month payroll month being processed
    * @param firstHours hours worked during the first cutoff
    * @param secondHours hours worked during the second cutoff
    * @param firstGross gross salary for the first cutoff
    * @param secondGross gross salary for the second cutoff
    * @param sssContribution SSS deduction amount
    * @param philhealthContribution PhilHealth deduction amount
    * @param pagibigContribution Pag-IBIG deduction amount
    * @param totalDeductions total deductions including tax
    * @param withholdingTax computed withholding tax
    * @param netSalary final salary after deductions
    */
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
            double withholdingTax,
            double netSalary) {

        String monthName = java.time.Month.of(month).name();
        int lastDayOfMonth = java.time.YearMonth.of(2024, month).lengthOfMonth();

        System.out.println("\n======================================");
        System.out.println("Employee Number : " + id);
        System.out.println("Employee Name   : " + employeeName);
        System.out.println("Birthday        : " + birthday);

        System.out.println("\nCutoff Date   : " + monthName + " 1 - 15");
        
        System.out.println("Hours Worked    : " + firstHours);
        System.out.println("Gross Salary    : " + firstGross);
        System.out.println("Net Salary      : " + firstGross);

        System.out.println("\nCutoff Date   : " + monthName + " 16 - " + lastDayOfMonth);

        System.out.println("Hours Worked    : " + secondHours);
        System.out.println("Gross Salary    : " + secondGross);
        
        System.out.println("\nDeductions");
        System.out.println("SSS             : " + sssContribution);
        System.out.println("PhilHealth      : " + philhealthContribution);
        System.out.println("Pag-IBIG        : " + pagibigContribution);
        System.out.println("Tax             : " + withholdingTax);
        
        System.out.println("Total Deduction : " + totalDeductions);
        System.out.println("Net Salary      : " + netSalary);
        System.out.println("======================================");
    }

    // ======================================================
    // ATTENDANCE PROCESSING AND WORK HOURS COMPUTATION
    // ======================================================
    
    /**
    * Generates the payroll report for a specific employee.
    *
    * This method reads attendance records from the attendance sheet
    * and calculates the total working hours for each payroll cutoff.
    * It then computes the gross salary, government deductions,
    * withholding tax, and final net salary for the employee.
    *
    * Payroll results are displayed using the formatted payroll
    * output method.
    *
    * @param employeeId the unique identifier of the employee
    * @param employeeName the full name of the employee
    * @param formattedBirthday the employee's formatted birth date
    * @param hourlyRate the employee's hourly pay rate
    * @param attendanceSheet the Excel sheet containing attendance records
    */
public static void generatePayrollForEmployee(
        int employeeId,
        String employeeName,
        String birthday,
        double hourlyRate,
        java.util.List<Row> employeeAttendance) {

    // ======================================================
    // PAYROLL COMPUTATION ALGORITHM
    // ======================================================
    // Algorithm Steps:
    // 1. Iterate through each month of the year.
    // 2. Read attendance records from the Excel attendance sheet.
    // 3. Filter attendance records belonging to the selected employee.
    // 4. Compute daily working hours based on login and logout times.
    // 5. Apply company policies:
    //      - Grace period for login
    //      - Official working hours (8:00 AM to 5:00 PM)
    //      - One-hour lunch deduction
    // 6. Separate working hours into two payroll cutoffs (1–15, 16–end).
    // 7. Compute gross salary for each cutoff.
    // 8. Calculate government deductions (SSS, PhilHealth, Pag-IBIG).
    // 9. Compute withholding tax based on taxable income.
    // 10. Display the payroll breakdown for the employee.
    
    for (int month = 1; month <= 12; month++) {

        Duration firstCutoff = Duration.ZERO;
        Duration secondCutoff = Duration.ZERO;

        for (Row row : employeeAttendance) {

            if (row.getRowNum() == 0) 
                continue; //skip column headers, to avoid reading text values
            
            //Gets the attendance date
            java.time.LocalDate attendanceDate =
                    row.getCell(3).getLocalDateTimeCellValue().toLocalDate();

            if (attendanceDate.getMonthValue() != month) continue;
            
            //Reads login time
            LocalTime login =
                    row.getCell(4).getLocalDateTimeCellValue().toLocalTime();

            LocalTime logout =
                    row.getCell(5).getLocalDateTimeCellValue().toLocalTime();

            // Apply grace period rule.
            // If the employee logs in before 8:00 AM or within the
            // allowed grace period (8:00–8:05 AM), the login time
            // is adjusted to the official start time of 8:00 AM
            if (login.isBefore(OFFICIAL_START) || !login.isAfter(GRACE_LIMIT)) {
                login = OFFICIAL_START;
            }
            
            //Limit Logout Time
            if (logout.isAfter(OFFICIAL_END)) {
                logout = OFFICIAL_END;
            }

            //Skip Invalid Time Records
            if (logout.isBefore(OFFICIAL_START) || login.isAfter(OFFICIAL_END)) {
                continue;
            }
            
            //Calculate Daily Work Duration
            Duration daily = Duration.between(login, logout);

            // Deduct the mandatory one-hour lunch break.
            // If the employee worked more than one hour,
            // subtract one hour from the computed work duration.
            if (daily.compareTo(Duration.ofHours(1)) > 0) {
                daily = daily.minusHours(1);
            } else {
                daily = Duration.ZERO;
            }
            
            //Assign Hours to Payroll Cutoff
            if (attendanceDate.getDayOfMonth() <= 15) {
                firstCutoff = firstCutoff.plus(daily);
            } else {
                secondCutoff = secondCutoff.plus(daily);
            }
        }

        //Convert Duration to Hours
        double firstHours = firstCutoff.toMinutes() / 60.0;
        double secondHours = secondCutoff.toMinutes() / 60.0;
        
        //Skip Months Without Attendance
        if (firstHours == 0 && secondHours == 0) continue;

        double firstGross = calculateGrossSalary(hourlyRate, firstHours);
        double secondGross = calculateGrossSalary(hourlyRate, secondHours);

        // ======================================================
        // DEDUCTION LOGIC (IMPORTANT)
        // ======================================================
        // Project requirement:
        // 1. Combine FIRST and SECOND cutoff gross → totalGross
        // 2. Compute ALL deductions based on totalGross
        // 3. Apply ALL deductions ONLY to SECOND cutoff salary
        
        double totalGross = firstGross + secondGross;
        
        // Government Contributions based on TOTAL gross
        double sssContribution = calculateSss(totalGross);
        double philhealthContribution = calculatePhilhealth(totalGross);
        double pagibigContribution = calculatePagibig(totalGross);

        double govDeductions = sssContribution + philhealthContribution + pagibigContribution;
        
        // Taxable income is based on total gross minus contributions
        double taxableIncome = totalGross - govDeductions;
        
        // Compute tax
        double tax = calculateWithholdingTax(taxableIncome);

        // Total deductions
        double totalDeductions = govDeductions + tax;
        
        // FINAL RULE: Deduct everything from SECOND cutoff only
        double netSalary = secondGross - totalDeductions;

        printFormattedPayroll(
                employeeId,
                employeeName,
                birthday,
                month,
                firstHours,
                secondHours,
                firstGross,
                secondGross,
                sssContribution,
                philhealthContribution,
                pagibigContribution,
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

           int inputId = getValidatedIntInput(scanner, "Enter Employee ID: ");

            boolean found = false;

            try {

                Workbook workbook = loadWorkbook("CP1_Grp9_MotorPH_Employee Data.xlsx");
                if (workbook == null) return;  // stop program if file failed

                Sheet employeeSheet = getSheet(workbook, 0);
                Sheet attendanceSheet = getSheet(workbook, 1);

                // ==========================================
                // GROUP ATTENDANCE ONCE
                // ==========================================
                java.util.Map<Integer, java.util.List<Row>> attendanceMap =
                        groupAttendanceByEmployee(attendanceSheet);

                DataFormatter formatter = new DataFormatter(); // To modify formats
                boolean firstRow = true; // skip header properly

                // =========================
                // SEARCH EMPLOYEE RECORD
                // =========================
                
                for (Row employeeRow : employeeSheet) {

                    if (firstRow) {
                        firstRow = false;
                        continue;
                    }

                    int id = Integer.parseInt(
                            formatter.formatCellValue(employeeRow.getCell(0)));

                    if (id == inputId) {

                        String lastName = formatter.formatCellValue(employeeRow.getCell(1));
                        String firstName = formatter.formatCellValue(employeeRow.getCell(2));
                        String employeeName = firstName + " " + lastName;

                        Cell birthdayCell = employeeRow.getCell(3);
                        Date birthday = birthdayCell.getDateCellValue();

                        SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
                        String formattedBirthday = sdf.format(birthday);

                        System.out.println("\n===== MotorPH Employee Details =====");
                        System.out.println("Employee Number : " + id);
                        System.out.println("Employee Name   : " + employeeName);
                        System.out.println("Birthday        : " + formattedBirthday);

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
                
                java.util.Map<Integer, java.util.List<Row>> attendanceMap =
                    groupAttendanceByEmployee(attendanceSheet);
                // =========================
                // MAIN MENU
                // =========================
                
                System.out.println("\n=== MotorPH Payroll System ===");
                System.out.println("1. Process Payroll");
                System.out.println("2. Exit");
                    
                    // VALIDATE MAIN MENU INPUT

                    int mainMenuOption = getValidatedIntInput(scanner, "Enter Option: ");
                    scanner.nextLine();

                // ======================================================
                // PROCESS PAYROLL
                // ======================================================
                
                if (mainMenuOption == 1) {

                    System.out.println("\n=== Process Payroll ===");
                    System.out.println("1. One Employee");
                    System.out.println("2. All Employees");
                    System.out.println("3. Exit");
                    
                int payrollOption;

                // VALIDATE SECOND MENU INPUT

                payrollOption = getValidatedIntInput(scanner, "Enter option: ");
                scanner.nextLine();

                    // =========================================
                    // ONE EMPLOYEE PAYROLL
                    // =========================================
                    
                    if (payrollOption == 1) {
                       
                       // VALIDATE EMPLOYEE ID INPUT

                        int inputId = getValidatedIntInput(scanner, "Enter Employee ID: ");

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

                              java.util.List<Row> records = attendanceMap.get(id);
                            
                                if (records != null) {
                                    generatePayrollForEmployee(
                                            id,
                                            employeeName,
                                            formattedBirthday,
                                            hourlyRate,
                                            records
                                    );
                                }

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
                    
                    else if (payrollOption == 2) {

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

                            java.util.List<Row> records = attendanceMap.get(id);

                            if (records != null) {
                                generatePayrollForEmployee(
                                        id,
                                        employeeName,
                                        formattedBirthday,
                                        hourlyRate,
                                        records
                                );
                            }
                        }
                    }
                

                    else if (payrollOption == 3) {
                        System.out.println("Exiting Process Payroll...");
                    }

                    else {
                        System.out.println("Invalid option.");
                    }

                }

                else if (mainMenuOption == 2) {
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