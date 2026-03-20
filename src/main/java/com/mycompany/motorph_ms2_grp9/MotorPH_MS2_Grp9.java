/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt 
to change this license
 */

// This tells which folder structure the class belongs to.
package com.mycompany.motorph_ms2_grp9; 

// ================================
// Java Standard Libraries
// ================================
// File handling
import java.io.FileInputStream; // Allows the program to Open and Read excel files.
import java.io.IOException;

// Date and formatting
import java.util.Date; // Stores date value such as employee birthdays.
import java.text.SimpleDateFormat; // Formats date values into a readable string format.

// User input
import java.util.Scanner; // Allows user input from the keyboard.

// =====================================
// Java Time API (Work Hour Calculation)
// =====================================
import java.time.LocalTime; //Stores time values such as log in and log out times.
import java.time.Duration; // Calculates the difference between two times.

// ====================================================
// Third-Party Libraries (Apache POI - Excel Handling)
// ====================================================
import org.apache.poi.ss.usermodel.*; // Access Excel components (Workbook, Sheet, Row, Cell).
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
 * DESIGN NOTE:
 * This program uses procedural decomposition instead of OOP.
 * Each method is responsible for a single task (calculation, formatting, 
 * or processing), improving readability, maintainability, and testability.
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
    
    // ================================
    // GOVERNMENT CONSTANTS
    // ================================
    // All constant values are based on Philippine government contribution tables (SSS, PhilHealth, Pag-IBIG)
    
    // PhilHealth
    private static final double PHILHEALTH_MIN_SALARY = 10000.0;
    private static final double PHILHEALTH_MAX_SALARY = 60000.0;
    private static final double PHILHEALTH_RATE = 0.03;

    // Pag-IBIG
    private static final double PAGIBIG_RATE_LOW = 0.01;
    private static final double PAGIBIG_RATE_HIGH = 0.02;
    private static final double PAGIBIG_MAX_CONTRIBUTION = 100.0;

    // SSS
    private static final int SSS_MIN_SALARY = 3250;
    private static final int SSS_MAX_SALARY = 24750;
    private static final double SSS_MIN_CONTRIBUTION = 135.0;
    private static final double SSS_MAX_CONTRIBUTION = 1125.0;
    private static final int SSS_STEP = 500;
    private static final double SSS_INCREMENT = 22.5;

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

        } catch (IOException e) {
            System.out.println("Error loading Excel file.");
            return null;
            }
    }
    
    // Provides access to a specific worksheet needed for employee or attendance data
    public static Sheet getSheet(Workbook workbook, int index) {
        if (workbook == null) return null;
            return workbook.getSheetAt(index);
    }
   
    // =============================================
    // ORGANIZES ATTENDANCE RECORDS BY EMPLOYEE ID
    // =============================================
    /**
    * PERFORMANCE IMPROVEMENT:
    * Instead of scanning the entire attendance sheet multiple times
    * (once per employee), this method processes the sheet ONLY ONCE
    * and stores records in a Map.
    *
    * This allows O(1) access to each employee’s attendance records,
    * significantly improving efficiency when processing payroll.
     * @param attendanceSheet
     * @return groupedData
    */
    public static java.util.Map<Integer, java.util.List<Row>>
        groupAttendanceByEmployee(Sheet attendanceSheet) {

    java.util.Map<Integer, java.util.List<Row>> groupedData 
        = new java.util.HashMap<>();

    for (Row row : attendanceSheet) {

        if (row.getRowNum() == 0) continue;

        int empId = (int) row.getCell(0).getNumericCellValue();

        groupedData.putIfAbsent(empId, new java.util.ArrayList<>());
        groupedData.get(empId).add(row);
    }

    return groupedData;
}

    // BUSINESS RULE IMPLEMENTATION:
    // 1. Employees who arrive on or before the grace period are treated as on-time.
    // 2. Work hours are capped between official start (8:00 AM) and end (5:00 PM).
    // 3. Invalid time entries (outside working hours) are ignored.
    // 4. A fixed 1-hour lunch break is deducted from total working time.
    public static Duration calculateDailyHours(LocalTime login, LocalTime logout) {

                // Skip invalid time records outside official working hours
        if (login.isBefore(OFFICIAL_START) || !login.isAfter(GRACE_LIMIT)) {
            login = OFFICIAL_START;
        }

        if (logout.isAfter(OFFICIAL_END)) {
            logout = OFFICIAL_END;
        }

        if (logout.isBefore(OFFICIAL_START) || login.isAfter(OFFICIAL_END)) {
            return Duration.ZERO;
        }

        Duration work = Duration.between(login, logout);

        if (work.compareTo(Duration.ofHours(1)) > 0) {
            return work.minusHours(1);
        }

        return Duration.ZERO;
    }

    public static void displayEmployeeDetails(int id, String name, String birthday) {
        System.out.println("\n===== MotorPH Employee Details =====");
        System.out.println("Employee Number : " + id);
        System.out.println("Employee Name   : " + name);
        System.out.println("Birthday        : " + birthday);
    }

    /**
    * Computes all payroll deductions based on total monthly gross income.
    *
    * IMPORTANT:
    * All government contributions (SSS, PhilHealth, Pag-IBIG) and tax
    * are calculated using the combined gross salary from both cutoffs.
    *
    * This ensures compliance with payroll rules where deductions are
    * based on total monthly earnings, not per cutoff.
     * @param totalGross the employee's combined monthly gross salary from both cutoffs
     * @return Deductions an array containing SSS, PhilHealth, Pag-IBIG, tax, and total deductions
    */
    public static double[] calculateDeductions(double totalGross) {

        double sss = calculateSss(totalGross);
        double philhealth = calculatePhilhealth(totalGross);
        double pagibig = calculatePagibig(totalGross);

        double totalGovernmentDeductions = sss + philhealth + pagibig;
        double taxableIncome = totalGross - totalGovernmentDeductions;
        double tax = calculateWithholdingTax(taxableIncome);

        double totalDeductions = totalGovernmentDeductions + tax;

        return new double[]{sss, philhealth, pagibig, tax, totalDeductions};
    }

    public static String[] getEmployeeDetails(Row row) {

        String lastName = row.getCell(1).getStringCellValue();
        String firstName = row.getCell(2).getStringCellValue();
        String fullName = firstName + " " + lastName;

        Date birthday = row.getCell(3).getDateCellValue();
        SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
        String formattedBirthday = sdf.format(birthday);

        double hourlyRate = row.getCell(18).getNumericCellValue();

        // Returns employee name, formatted birthday, and hourly rate in a fixed order.
        // [0]=name, [1]=birthday, [2]=hourlyRate
        return new String[]{fullName, formattedBirthday, String.valueOf(hourlyRate)};
    }

    // ======================================================
    // CALCULATIONS
    // ======================================================
    
    // This calculation is based on a fixed hourly pay structure and
    // does not include overtime pay, bonuses, or allowances.
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

        int start = SSS_MIN_SALARY;
        int end = SSS_MAX_SALARY;
        double minContribution = SSS_MIN_CONTRIBUTION;
        double maxContribution = SSS_MAX_CONTRIBUTION;
        
        // The SSS contribution increases in fixed salary brackets.
        // This loop simulates the contribution table by incrementing
        // the salary range and contribution amount step-by-step
        // until it matches the employee's salary bracket.
        int startRange = 3250;
        int endRange = 3750;
        int computeRangeAdder = SSS_STEP;
        
        double finalContribution = 157.50;
        double contributionAdder = SSS_INCREMENT;
        
        if (totalGross < start) {
            return minContribution;
            
        } else if (totalGross >= end) {
            return maxContribution;
            
        } else if (totalGross >= start && totalGross < end) {
            
            do {
                
                startRange = startRange + computeRangeAdder;
                endRange = endRange + computeRangeAdder;
                finalContribution = finalContribution + contributionAdder;
                
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


        double salaryBase;

        if (totalGross < PHILHEALTH_MIN_SALARY) {
            salaryBase = PHILHEALTH_MIN_SALARY;

        } else if (totalGross > PHILHEALTH_MAX_SALARY) {
            salaryBase = PHILHEALTH_MAX_SALARY;

        } else {
            salaryBase = totalGross;
        }

        double totalPremium = salaryBase * PHILHEALTH_RATE;
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
            employeeRate = PAGIBIG_RATE_LOW;

        } else if (totalGross > 1500) {
            employeeRate = PAGIBIG_RATE_HIGH;

        } else {
            return 0;
        }

        double pagIbigContribution = totalGross * employeeRate;

        if (pagIbigContribution > PAGIBIG_MAX_CONTRIBUTION) {
            pagIbigContribution = PAGIBIG_MAX_CONTRIBUTION;
        }

        return pagIbigContribution;
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
    * @param birthday the employee's formatted birth date
    * @param hourlyRate the employee's hourly pay rate
    * @param employeeAttendance
    */
    public static void generatePayrollForEmployee(
        int employeeId,
        String employeeName,
        String birthday,
        double hourlyRate,
        java.util.List<Row> employeeAttendance) {

    // ======================================================
    // PAYROLL ALGORITHM OVERVIEW
    // ======================================================

    // This method computes payroll by processing pre-grouped attendance records.
    // For each month:
    // 1. Total work hours are computed per cutoff (1–15, 16–end).
    // 2. Gross pay is calculated separately for each cutoff.
    // 3. Both cutoffs are combined to determine total monthly income.
    // 4. Government deductions and tax are computed using total income.
    // 5. All deductions are applied ONLY to the second cutoff salary,
    //    based on project rules.
    //
    // This approach ensures both accuracy and compliance with payroll policies.
    
    for (int month = 1; month <= 12; month++) {

        Duration firstCutoff = Duration.ZERO;
        Duration secondCutoff = Duration.ZERO;
        
        // NOTE: We are NOT scanning the entire attendance sheet here.
        // We are only iterating through PRE-GROUPED records for this employee.
        for (Row row : employeeAttendance) {

            if (row.getRowNum() == 0) 
                continue; //skip column headers, to avoid reading text values
            
            // Retrieves the attendance date and converts it to LocalDate
            // for use in payroll cutoff and monthly processing.
            java.time.LocalDate attendanceDate =
                    row.getCell(3).getLocalDateTimeCellValue().toLocalDate();

            if (attendanceDate.getMonthValue() != month) continue;
            
            // Extracts the employee's login time from the attendance record.
            // The value is read from the Excel cell as a date-time and converted
            // to LocalTime to enable accurate time-based calculations.
            LocalTime login =
                    row.getCell(4).getLocalDateTimeCellValue().toLocalTime();

            LocalTime logout =
                    row.getCell(5).getLocalDateTimeCellValue().toLocalTime();
            
            // Compute adjusted working hours based on company policies
            Duration daily = calculateDailyHours(login, logout);
            
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
        // FINAL PAYROLL RULE IMPLEMENTATION
        // ======================================================
        
        // Project requirements:
        // STEP 1: Combine BOTH cutoffs to get TOTAL monthly gross
        double totalGross = firstGross + secondGross;
        
        // STEP 2: Compute ALL deductions based on TOTAL gross
        // (SSS, PhilHealth, Pag-IBIG, Tax are based on monthly income)
        double[] deductions = calculateDeductions(totalGross);

        // STEP 3: Apply ALL deductions ONLY to SECOND cutoff
        // (Note: deductions are not split across cutoffs)
        double sssContribution = deductions[0];
        double philhealthContribution = deductions[1];
        double pagibigContribution = deductions[2];
        double tax = deductions[3];
        double totalDeductions = deductions[4];
        
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
        System.out.print("Username: ");
            String inputUsername = scanner.nextLine();

        System.out.print("Password: ");
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

                try (Workbook workbook = loadWorkbook("CP1_Grp9_MotorPH_Employee Data.xlsx")) {
                    if (workbook == null) return;  // stop program if file failed
                    
                    Sheet employeeSheet = getSheet(workbook, 0);
                    Sheet attendanceSheet = getSheet(workbook, 1);
                    
                    // ==========================================
                    // GROUP ATTENDANCE ONCE
                    // ==========================================
                    
                    // Pre-process attendance data once and group by employee
                    // This avoids repeated scanning of the Excel file
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
                            
                            String[] employeeDetails = getEmployeeDetails(employeeRow);
                            
                            String employeeName = employeeDetails[0];
                            String formattedBirthday = employeeDetails[1];
                            
                            displayEmployeeDetails(id, employeeName, formattedBirthday);
                            
                            found = true;
                            break;
                        }
                    }
                    
                    if (!found) {
                        System.out.println("Employee number does not exist");            
                    }
                }

            } catch (IOException | NumberFormatException e) {
                System.out.println("An error occurred while processing data.");
            }
        }
  
        // ======================================================
        // PAYROLL STAFF LOGIN
        // ======================================================

        else if (inputUsername.equals("payroll_staff") 
                && inputPassword.equals("12345")) {

            System.out.println("\nLogin Successful!");

            try {

                try (Workbook workbook = loadWorkbook
                    ("CP1_Grp9_MotorPH_Employee Data.xlsx")) {
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
                    switch (mainMenuOption) {
                        case 1 -> {
                            System.out.println("\n=== Process Payroll ===");
                            System.out.println("1. One Employee");
                            System.out.println("2. All Employees");
                            System.out.println("3. Exit");
                            int payrollOption;
                            
                            payrollOption = getValidatedIntInput(scanner, "Enter option: ");
                            scanner.nextLine();
                            // =========================================
                            // ONE EMPLOYEE PAYROLL
                            // =========================================
                            switch (payrollOption) {
                                case 1 -> {
                                    int inputId = getValidatedIntInput(scanner, "Enter Employee ID: ");
                                    boolean found = false;
                                    for (Row row : employeeSheet) {
                                        
                                        if (row.getRowNum() == 0) continue;
                                        
                                        int id = (int) row.getCell(0)
                                                .getNumericCellValue();
                                        
                                        if (id == inputId) {
                                            
                                            String[] details = getEmployeeDetails(row);
                                            
                                            String employeeName = details[0];
                                            String formattedBirthday = details[1];
                                            double hourlyRate = Double.parseDouble(details[2]);
                                            
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
                                case 2 -> {
                                    for (Row row : employeeSheet) {
                                        
                                        if (row.getRowNum() == 0) continue;
                                        
                                        int id = (int) row.getCell(0).getNumericCellValue();
                                        
                                        String[] details = getEmployeeDetails(row);
                                        
                                        String employeeName = details[0];
                                        String formattedBirthday = details[1];
                                        double hourlyRate = Double.parseDouble(details[2]);
                                        
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
                                case 3 -> System.out.println("Exiting Process Payroll...");
                                default -> System.out.println("Invalid option.");
                            }
                        }
                        case 2 -> System.out.println("Exiting program...");
                        default -> System.out.println("Invalid option.");
                    }
                }

                } 
                catch (IOException | NumberFormatException e) {
                    System.out.println("An error occurred while processing data.");
                }
                
        }
        else {
            System.out.println("Invalid username or password.");
        }
      }
}
