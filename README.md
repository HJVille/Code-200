# MotorPH Payroll System

**Course:** MO-IT101 - Computer Programming 1  
**Submission Purpose:** Terminal Assessment: Final Code Submission  
**Group:** Group 9  
**Submission Branch:** `Terminal_Assessment_Final_Code_Submission`

---

## Team Details

| Member | Contribution |
|------|------|
| **Harty Joy Villegas** | Assisted in analyzing the system requirements based on the assignment instructions, managed the GitHub repository and version control, organized project files, maintained repository updates, and prepared the project documentation including this README. Helped review the system implementation against the required specifications and participated in project discussions and coordination. |
| **Kimberly Joy Goyena** | Developed and finalized the working payroll system implementation, structured the Excel dataset, and implemented the payroll computation logic for hours worked, gross salary, deductions, and net salary. Led testing, debugging, validation of payroll outputs, and refinement of the code structure and logic. |

**Note:**  
This submission reflects the collaboration of **Harty Joy Villegas** and **Kimberly Joy Goyena** in system development, testing, documentation, and repository preparation.

---

## Program Details

The **MotorPH Payroll System** is a Java console-based payroll application that reads employee information and attendance records from an Excel workbook and computes payroll results based on hours worked and hourly rate.

The system was developed as a simplified academic payroll processing program. It demonstrates:

- file handling using an Excel dataset
- console-based user interaction
- employee detail lookup
- attendance record processing
- payroll cutoff computation
- gross salary computation
- government deduction computation
- withholding tax computation
- payroll summary presentation

The program uses the **Apache POI** library to read and process Excel files.

---

## What the System Does

The program supports two user roles:

- **employee**
- **payroll_staff**

If the user logs in as **employee**, the program allows employee detail lookup using an employee ID.

If the user logs in as **payroll_staff**, the program allows payroll processing for:

- one employee
- all employees

The system reads data from the Excel file:

```text
CP1_Grp9_MotorPH_Employee Data.xlsx
```

The workbook contains two sheets:

- `Employee Details`
- `Attendance Record`

---

## Login Credentials

### Valid Usernames

```text
employee
payroll_staff
```

### Password

```text
12345
```

If the login credentials are incorrect, the system displays:

```text
Invalid username or password.
```

---

## Program Execution Flow

```text
Start Program
      |
      v
Display Welcome Screen
      |
      v
Ask for Username and Password
      |
      v
Echo Entered Username and Password
      |
      v
Validate Credentials
      |
      +-----------------------------+
      |                             |
      v                             v
Employee Login                 Payroll Staff Login
      |                             |
      v                             v
Login Successful               Login Successful
      |                             |
      v                             v
Enter Employee ID              Load Workbook
      |                             |
      v                             v
Load Workbook                  Show Main Menu
      |                             |
      v                             v
Display Employee Info          Process One or All Employees
                                    |
                                    v
                             Compute Payroll Per Employee
                                    |
                                    v
                             Display Payroll Summary
```

---

## Actual Console Behavior

After the user enters the username and password, the program echoes the entered credentials in the console and prints:

```text
=== Log In as Employee ===
Username: <entered username>
Password: <entered password>
```

This output appears before the program branches into the employee login or payroll staff login flow.

When the project is run in NetBeans, the Output window may also show Maven execution lines such as build messages and `BUILD SUCCESS`.

The console may also display a Log4j status warning while reading the workbook. These messages appear in the run output, but they are separate from the payroll computation itself.

---

## Employee Access

If the user logs in as **employee**, the system:

1. asks for an employee ID
2. loads the Excel workbook
3. reads the `Employee Details` sheet
4. searches for the matching employee ID
5. displays:
   - employee number
   - employee name
   - birthday

The employee details section is displayed with the heading:

```text
===== MotorPH Employee Details =====
```

If the employee number does not exist, the system displays:

```text
Employee number does not exist
```

The employee login branch is for **viewing employee details only**. It does not compute payroll.

---

## Payroll Staff Access

If the user logs in as **payroll_staff**, the system displays:

```text
1. Process Payroll
2. Exit
```

If **Process Payroll** is selected, the system then displays:

```text
1. One Employee
2. All Employees
3. Exit
```

If an invalid menu option is entered, the system displays:

```text
Invalid option.
```

---

## One Employee Payroll

The payroll staff enters an employee ID.

If the employee exists in the dataset, the system:

- retrieves employee information
- retrieves attendance records for that employee
- computes payroll for months that contain attendance data
- displays the payroll summary in the console

If the employee ID is not found in this payroll branch, the system displays:

```text
Employee ID Not Found.
```

---

## All Employees Payroll

If **All Employees** is selected, the system loops through all employee records in the employee sheet and computes payroll for every employee who has attendance records.

---

## Payroll Rules Implemented

The system follows the payroll rules implemented in the code.

### Working Hours

Only working hours within the following schedule are counted:

```text
8:00 AM to 5:00 PM
```

The program also uses a grace period:

```text
8:05 AM
```

### Attendance Adjustment Rules

- If login is before `8:00 AM`, login is adjusted to `8:00 AM`
- If login is between `8:00 AM` and `8:05 AM`, login is also adjusted to `8:00 AM`
- If logout is after `5:00 PM`, logout is adjusted to `5:00 PM`
- If a record falls outside valid work hours, the record is skipped

### Lunch Deduction

If the employee worked for more than 1 hour, the program deducts **1 hour** for lunch.

### Payroll Cutoff

The system divides attendance into two cutoffs:

- **First cutoff:** day `1` to `15`
- **Second cutoff:** day `16` to end of month

### Gross Salary Formula

```text
Gross Salary = Hourly Rate x Total Hours Worked
```

---

## Government Deductions

Before computing deductions, the program combines the two payroll cutoffs:

```text
Total Gross Salary = First Cutoff Gross + Second Cutoff Gross
```

The following deductions are calculated:

- SSS Contribution
- PhilHealth Contribution
- Pag-IBIG Contribution
- Withholding Tax

The deduction flow used by the system is:

```text
Government Deductions = SSS + PhilHealth + Pag-IBIG
Taxable Income = Total Gross Salary - Government Deductions
Total Deductions = Government Deductions + Withholding Tax
Net Salary = Second Cutoff Gross - Total Deductions
```

All deductions are computed based on the **combined monthly gross salary**, but they are applied to the **second cutoff only**.

Because of this:

- the first cutoff is displayed as gross salary and net salary without deductions applied there
- the second cutoff carries the deduction load for the month

---

## Output Produced by the System

For payroll processing, the system prints:

- employee number
- employee name
- birthday
- first cutoff date
- first cutoff hours worked
- first cutoff gross salary
- first cutoff displayed net salary
- second cutoff date
- second cutoff hours worked
- second cutoff gross salary
- SSS
- PhilHealth
- Pag-IBIG
- tax
- total deduction
- final net salary

The program displays payroll results directly in the console.

---

## Data Source

The system reads employee and attendance data from:

```text
CP1_Grp9_MotorPH_Employee Data.xlsx
```

### Sheet 1: Employee Details

This sheet stores data such as:

- employee ID
- last name
- first name
- birthday
- government IDs
- salary information
- hourly rate

### Sheet 2: Attendance Record

This sheet stores:

- employee ID
- last name
- first name
- date
- log in time
- log out time

The current implementation depends on the expected sheet order and required column positions in the workbook.

---

## How to Run the System

### Run in NetBeans

1. Open NetBeans
2. Open the project as a Maven project
3. Make sure the project uses the appropriate JDK
4. Run the main class:
   `com.mycompany.motorph_ms2_grp9.MotorPH_MS2_Grp9`

### Required Local Files

To run successfully, the following files must be present locally:

- `pom.xml`
- `src/main/java/com/mycompany/motorph_ms2_grp9/MotorPH_MS2_Grp9.java`
- `CP1_Grp9_MotorPH_Employee Data.xlsx`

The Excel file is loaded by filename from the project directory. If the file is renamed, moved, or removed, the workbook will not load correctly.

---

## Testing Guide

The following test scenarios can be used to verify the current system behavior.

### Test 1: Invalid Login

**Input**

- invalid username
- invalid password

**Expected Result**

```text
Invalid username or password.
```

### Test 2: Employee Login with Valid Employee ID

**Input**

- username: `employee`
- password: `12345`
- employee ID: existing employee ID from the dataset

**Expected Result**

The system displays:

- employee number
- employee name
- birthday

### Test 3: Employee Login with Invalid Employee ID

**Input**

- username: `employee`
- password: `12345`
- employee ID: non-existing employee ID

**Expected Result**

```text
Employee number does not exist
```

### Test 4: Payroll Staff Login, One Employee

**Input**

- username: `payroll_staff`
- password: `12345`
- option: `1`
- payroll option: `1`
- employee ID: existing employee ID

**Expected Result**

The system prints the payroll summary for the chosen employee for months that contain attendance records.

### Test 5: Payroll Staff Login, All Employees

**Input**

- username: `payroll_staff`
- password: `12345`
- option: `1`
- payroll option: `2`

**Expected Result**

The system prints payroll summaries for all employees that have attendance records.

### Test 6: Invalid Numeric Menu or ID Input

**Input**

- letters or invalid non-numeric input when an integer is expected

**Expected Result**

```text
Invalid input. Please enter a number.
```

### Test 7: Payroll Staff Login with Non-existing Employee ID

**Input**

- username: `payroll_staff`
- password: `12345`
- option: `1`
- payroll option: `1`
- employee ID: non-existing employee ID

**Expected Result**

```text
Employee ID Not Found.
```

### Testing Checklist

During testing, confirm that:

- the program starts and displays the welcome banner
- valid credentials allow access to the correct login role
- invalid credentials are rejected
- employee details can be retrieved from the workbook
- payroll staff can process payroll for one employee
- payroll staff can process payroll for all employees
- deductions are displayed in the second cutoff section
- invalid integer input does not crash the program
- invalid menu options display the proper error message
- payroll one-employee processing displays `Employee ID Not Found.` for non-existing IDs

---

## Repository Structure

```text
MO-IT101-Group9
|
|-- src
|   `-- main
|       `-- java
|           `-- com
|               `-- mycompany
|                   `-- motorph_ms2_grp9
|                       `-- MotorPH_MS2_Grp9.java
|
|-- CP1_Grp9_MotorPH_Employee Data.xlsx
|-- pom.xml
`-- README.md
```

### Meaning of the Main Files

- `MotorPH_MS2_Grp9.java`  
  Main Java source file containing the full program logic.

- `CP1_Grp9_MotorPH_Employee Data.xlsx`  
  Excel workbook used as the system data source.

- `pom.xml`  
  Maven configuration file that defines the project settings and dependencies.

---

## Current Status

This repository represents the **final submission version** of the MotorPH Payroll System for terminal assessment submission.

### Currently Implemented Features

- Excel file reading using Apache POI
- Employee detail lookup
- Attendance record processing
- Payroll processing for one employee
- Payroll processing for all employees
- Hours worked computation based on attendance logs
- Gross salary computation using hourly rate
- Government deduction computation
- Withholding tax computation
- Console-based payroll summary output
- Input validation for numeric entries

### Current Implementation Notes

- the system uses a fixed Excel workbook as its data source
- payroll processing is performed in a console environment
- employee and payroll staff access are handled through hardcoded login credentials
- the main business logic is implemented in a single Java source file

---

## Project Summary

The **MotorPH Payroll System** is a Java-based academic payroll application that demonstrates file handling, input validation, attendance processing, payroll computation, deduction computation, and console-based reporting.

It is centered around one Java implementation file, one Excel workbook, and one Maven configuration file. The system allows:

- employee detail lookup
- payroll processing for one employee
- payroll processing for all employees
- payroll calculation based on attendance and hourly rate
- deduction calculation using combined monthly gross salary

This repository branch contains the working final implementation of the project for terminal assessment submission.
