# MotorPH Payroll System

**Course:** MO-IT101 - Computer Programming 1  
**Assessment:** Terminal Assessment  
**Group 9**

---

# Team Details

| Member | Contribution |
|------|------|
| **Harty Joy Villegas** | Developed the initial versions of the payroll program and assisted in analyzing the system requirements based on the assignment instructions. Managed the GitHub repository and version control, organized project files, and maintained system updates. Prepared and maintained the project documentation including the README, ensuring that the documentation aligned with the project requirements and mentor feedback. Coordinated project progress, reviewed system implementation against the required specifications, and actively participated in team discussions, meetings, and mentoring sessions. |
| **Kimberly Joy Goyena** | Developed and finalized the working payroll system implementation used in the project, including structuring the Excel dataset and implementing the payroll computation logic for hours worked, gross salary, deductions, and net salary. Worked closely with Harty throughout the development process to ensure the system followed the assignment requirements. Led system testing and revision, including debugging payroll computations, validating outputs, and improving code readability and error handling. Also helped organize team meetings and ensured consistent progress updates during the project development. |

**Note:**  
This repository reflects the collaboration of **Harty Joy Villegas** and **Kimberly Joy Goyena** in system development, testing, documentation, and repository preparation.

**Project Plan:**  
[MotorPH Project Plan](https://docs.google.com/spreadsheets/d/107AtUMdNt3sHL_d9LgfsAUAUdq5YrEG62ZbGcSEGtmE/edit?usp=sharing)

---

# Program Details

The **MotorPH Payroll System** is a Java console-based application that simulates a simplified payroll processing system.

The program reads employee information and attendance records from an **Excel dataset** and calculates payroll results based on the number of hours worked during payroll cutoffs.

The system performs the following functions:

- Reads employee information from an Excel file
- Processes attendance records
- Calculates total hours worked
- Computes gross salary based on hourly rate
- Applies government deductions
- Calculates net salary
- Displays payroll summaries in the console

The program uses the **Apache POI** library to read and process Excel files containing employee and attendance data.

---

# Login Credentials

**Valid usernames**

```text
employee
payroll_staff
```

**Password**

```text
12345
```

If the login credentials are incorrect, the program displays:

```text
Invalid username or password.
```

and the program terminates.

---

# Program Execution Flow

```text
Start Program
      ↓
Display Welcome Banner
      ↓
Ask for Username and Password
      ↓
Echo Entered Username and Password
      ↓
Validate Credentials
      ↓
Proceed Based on Username
```

---

# Employee Access

If the user logs in as **employee**, the system displays:

```text
1. Enter your employee number
2. Exit the program
```

If the user selects the employee-number option, the system asks for:

```text
Enter employee number:
```

When a valid employee number is entered, the system displays:

- Employee Number
- Employee Name
- Birthday

If the employee number does not exist:

```text
Employee number does not exist.
```

The employee menu remains active until **Exit the program** is selected.

---

# Payroll Staff Access

If the user logs in as **payroll_staff**, the system displays:

```text
1. Process Payroll
2. Exit the program
```

If **Process Payroll** is selected, the system then displays:

```text
1. One Employee
2. All Employees
3. Exit the program
```

If an invalid menu option is entered, the system displays:

```text
Invalid option.
```

The payroll menus remain active until **Exit the program** is selected.

---

# One Employee Payroll

The payroll staff enters an **employee number**.

If the employee exists in the dataset, the system retrieves the employee details and computes payroll information using the attendance records grouped for that employee.

Payroll records are displayed for the months that contain attendance data in the workbook.

### First Cutoff

The first cutoff section displays:

- Cutoff Date
- Hours Worked
- Gross Salary
- Net Salary

### Second Cutoff

The second cutoff section displays:

- Cutoff Date
- Hours Worked
- Gross Salary

### Deductions

The deduction section displays:

- SSS
- PhilHealth
- Pag-IBIG
- Tax
- Total Deductions
- Net Salary

Government deductions are applied during the **second payroll cutoff**.

---

# All Employees Payroll

If **All Employees** is selected, the program loops through the employee sheet and generates payroll output for employees that have attendance records available in the grouped attendance map.

---

# Payroll Rules Implemented

The system follows payroll rules based on the assignment instructions and mentor clarifications.

## Working Hours

Only working hours within the following schedule are counted:

```text
8:00 AM to 5:00 PM
```

The system also applies a grace period of:

```text
8:05 AM
```

Example calculations:

| Login | Logout | Hours Worked |
|------|------|------|
| 8:30 AM | 5:30 PM | 7.5 hours |
| 8:05 AM | 5:00 PM | 8 hours |
| 8:05 AM | 4:30 PM | 7.5 hours |

Extra hours outside the official working schedule are **not included** in payroll computation.

## Government Deductions

Before computing deductions, the program combines the two payroll cutoffs.

```text
Total Gross Salary = First Cutoff Gross + Second Cutoff Gross
```

The following deductions are calculated:

- SSS Contribution
- PhilHealth Contribution
- Pag-IBIG Contribution
- Withholding Tax

Net salary is computed after subtracting total deductions from the **second cutoff gross salary**.

## Important Note on Decimal Values

Salary values and hours worked may display multiple decimal places. The system does **not round off** values.

---

# Data Source

The system reads employee and attendance data from the Excel dataset:

```text
CP1_Grp9_MotorPH_Employee Data.xlsx
```

Dataset contents:

| Data | Description |
|------|------|
| Employee ID | Unique employee number |
| Employee Name | Full name of employee |
| Birthday | Employee birthdate |
| Hourly Rate | Salary per hour |
| Attendance Records | Login and logout times |

The program reads this dataset using **Apache POI**.

---

# Repository Structure

```text
MO-IT101-Group9
│
├── src
│   └── main
│       └── java
│           └── com
│               └── mycompany
│                   └── motorph_ms2_grp9
│                       └── MotorPH_MS2_Grp9.java
│
├── CP1_Grp9_MotorPH_Employee Data.xlsx
├── pom.xml
└── README.md
```

---

# Sample Program Output

### Example Login

```text
==================================
       WELCOME TO MOTORPH
==================================

Username: payroll_staff
Password: 12345

=== Login ===
Username: payroll_staff
Password: 12345
```

### Example Payroll Menu

```text
=== MotorPH Payroll System ===

1. Process Payroll
2. Exit the program
Enter an option: 1

=== Process Payroll ===
1. One Employee
2. All Employees
3. Exit the program
Enter an option: 1
Enter employee number: 10001
```

### Example Payroll Output

```text
======================================
Employee Number : 10001
Employee Name   : John Doe
Birthday        : 05/15/1995

Cutoff Date     : June 1 - 15
Hours Worked    : 80
Gross Salary    : 8000
Net Salary      : 8000

Cutoff Date     : June 16 - 30
Hours Worked    : 80
Gross Salary    : 8000

Deductions
SSS             : 500
PhilHealth      : 240
Pag-IBIG        : 100
Tax             : 450
Total Deductions: 1290
Net Salary      : 6710
======================================
```

---

# Internal Testing and Revision

Testing and revision were handled through the collaboration of **Kimberly Joy Goyena** and **Harty Joy Villegas**.

Testing activities included:

- validating payroll computations against the attendance records
- checking employee detail lookup
- checking payroll processing for one employee
- checking payroll processing for all employees
- checking invalid login behavior
- checking numeric input validation
- reviewing output wording and menu behavior

These activities helped ensure the correctness and stability of the payroll computation before final cleanup and submission preparation.

---

# External Testing

As advised during the project process, **Team 10 - IceCreamCode** was added to the repository as external collaborators for independent QA checking.

The team provided an external test document titled:

```text
Team10_IceCreamCode_QA_Test_Document - Google Sheets.pdf
```

The submitted QA sheet covered multiple areas of the system, including:

- login functionality
- database and file loading
- navigation and menu flow
- employee detail lookup
- one-employee and all-employees payroll processing
- government deduction calculations
- payroll output printing
- attendance-based working-hours rules

Examples of the documented external test cases include:

- launching the application
- logging in with valid `employee` credentials
- logging in with valid `payroll_staff` credentials
- handling invalid login credentials
- successfully reading the project data file
- viewing the employee and payroll staff menu options
- processing payroll for one employee and for all employees
- verifying gross salary, net salary, SSS, PhilHealth, Pag-IBIG, and income tax calculations
- verifying that calculations and printed values were not rounded off
- checking standard working hours, grace period handling, late login, overtime, and early logout behavior

Most of the documented external test cases were marked **Pass** in the submitted sheet.

The QA document also recorded specific notes during testing:

- a **Log4j StatusLogger** warning appeared in the console during some test runs because of a missing logging implementation
- one menu-related test case for viewing the employee menu options was marked **Fail** during their run, and that feedback helped guide the later adjustment of the employee menu flow in the repository

The Log4j warning was environment-related and has since been addressed in the project configuration by adding the required Log4j runtime dependency in `pom.xml`.

This external testing helped confirm that the system was usable by another team outside the original developers and that the login flow, file loading, and basic program access worked as expected.

---

# Current Status

This repository currently represents the **Terminal Assessment** version of the MotorPH Payroll System.

Currently implemented features:

- Excel file reading
- Attendance record processing
- Salary computation
- Government deduction computation
- Console-based employee detail display
- Console-based payroll output
- Numeric input validation
- Menu loops that remain active until exit is selected

Current implementation notes:

- the project uses one Java source file
- the project does not use OOP concepts
- the current workbook is read-only and is not modified by the program
- the system uses grouped attendance records for payroll processing

---

# Project Summary

The **MotorPH Payroll System** is a Java-based academic payroll application that demonstrates file handling, input validation, attendance processing, payroll computation, deduction computation, and console-based reporting.

This repository contains the current working version of the MotorPH Payroll System for the Terminal Assessment. It reflects the implemented features, testing updates, documentation revisions, and QA-based improvements completed for this stage of the project.
