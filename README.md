# MotorPH Payroll System

**Course:** MO-IT101 – Computer Programming 1  
**Milestone:** Milestone 2 – Initial Code Implementation  
**Group 9** 

---

# Team Details

| Member | Contribution |
|------|------|
| **Harty Joy Villegas** | Developed the initial versions of the payroll program and assisted in analyzing the system requirements based on the assignment instructions. Managed the GitHub repository and version control, organized project files, and maintained system updates. Prepared and maintained the project documentation including the README, ensuring that the documentation aligned with the project requirements and mentor feedback. Coordinated project progress, reviewed system implementation against the required specifications, and actively participated in team discussions, meetings, and mentoring sessions. |
| **Kimberly Joy Goyena** | Developed and finalized the working payroll system implementation used in the project, including structuring the Excel dataset and implementing the payroll computation logic for hours worked, gross salary, deductions, and net salary. Worked closely with Harty throughout the development process to ensure the system followed the assignment requirements. Led system testing and revision, including debugging payroll computations, validating outputs, and improving code readability and error handling. Also helped organize team meetings and ensured consistent progress updates during the project development. || **Charmel Sammah** | Attended the first meeting after Milestone 1 and presented an initial code attempt. However, there was no further participation in system development, project updates, or coordination afterward. |
| **Charmel Sammah** | Attended the first meeting after Milestone 1 and presented an initial code attempt. However, there was no further participation in system development, project updates, or coordination afterward. |
| **Sophia Reign Pagatpatan** | Participated in early discussions and presented initial code attempts. However, there was no further participation in later stages of system development, updates, or coordination with the group. |

**Note:**  
Most of the system development, documentation, and repository updates were completed through the collaboration of **Harty Joy Villegas** and **Kimberly Joy Goyena**.

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

The program uses the **Apache POI library** to read and process Excel files containing employee and attendance data.

---

# Login Credentials

**Valid usernames**

```
employee
payroll_staff
```

**Password**

```
12345
```

If the login credentials are incorrect, the program displays:

```
Invalid username or password
```

and the program terminates.

---

# Program Execution Flow

```
Start Program
      ↓
User Login
      ↓
Validate Credentials
      ↓
Load Excel Workbook
      ↓
Retrieve Employee Information
      ↓
Process Attendance Records
      ↓
Calculate Hours Worked
      ↓
Compute Gross Salary
      ↓
Apply Government Deductions
      ↓
Calculate Net Salary
      ↓
Display Payroll Summary
```

---

# Employee Access

If the user logs in as **employee**, the system displays:

```
1 Enter Employee Number
2 Exit Program
```

When a valid employee number is entered, the system displays:

- Employee Number
- Employee Name
- Birthday

If the employee number does not exist:

```
Employee number does not exist
```

---

# Payroll Staff Access

If the user logs in as **payroll_staff**, the system displays:

```
1 Process Payroll
2 Exit Program
```

If **Process Payroll** is selected:

```
1 One Employee
2 All Employees
3 Exit
```

---

# One Employee Payroll

The payroll staff enters an **employee ID**.

If the employee exists in the dataset, the system retrieves the employee details and computes payroll information.

Payroll records are displayed for the period **June to December**.

---

## First Cutoff

```
Cutoff Date: June 1 to June 15
Total Hours Worked
Gross Salary
Net Salary
```

---

## Second Cutoff

```
Cutoff Date: June 16 to June 30
Total Hours Worked
Gross Salary
```

Each Deduction

```
SSS
PhilHealth
Pag-IBIG
Tax
```

Total Deductions  
Net Salary

Government deductions are applied during the **second payroll cutoff**.

---

# Payroll Rules Implemented

The system follows payroll rules provided in the assignment instructions.

## Working Hours

Only working hours within the following schedule are counted:

```
8:00 AM to 5:00 PM
```

Example calculations:

| Login | Logout | Hours Worked |
|------|------|------|
| 8:30 AM | 5:30 PM | 7.5 hours |
| 8:05 AM | 5:00 PM | 8 hours |
| 8:05 AM | 4:30 PM | 7.5 hours |

Extra hours outside the official working schedule are **not included in payroll computation**.

---

# Government Deductions

Before computing deductions, the program combines the two payroll cutoffs.

```
Total Gross Salary = First Cutoff + Second Cutoff
```

The following deductions are calculated:

- SSS Contribution
- PhilHealth Contribution
- Pag-IBIG Contribution
- Withholding Tax

Net salary is computed after subtracting total deductions.

---

# Data Source

The system reads employee and attendance data from the Excel dataset:

```
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

```
MotorPH Payroll System
│
├── src
│   └── com.mycompany.motorph_ms2_grp9
│        └── MotorPH_MS2_Grp9.java
│
├── resources
│   └── CP1_Grp9_MotorPH_Employee Data.xlsx
│
└── README.md
```

---

# Sample Program Output

Example login

```
==================================
        WELCOME TO MOTORPH
==================================

Username: payroll_staff
Password: 12345
```

Example payroll output

```
======================================
Employee Number: 10001
Employee Name: John Doe
Birthday: 05/15/1995

Cutoff Date: JUNE 1 to JUNE 15
Total Hours Worked: 80
Gross Salary: 8000
Net Salary: 8000

Cutoff Date: JUNE 16 to JUNE 30
Total Hours Worked: 80
Gross Salary: 8000

SSS: 500
PhilHealth: 240
Pag-IBIG: 100
Tax: 450

Total Deductions: 1290
Net Salary: 6710
======================================
```

---

# Project Plan

The project plan outlines the development schedule, task assignments, and estimated effort required to complete the payroll system.

Development activities include:

- requirements visualization
- employee details implementation
- hours worked calculation
- gross wage computation
- net wage computation
- debugging and testing
- system documentation

**Estimated effort**

```
Total Effort: 47 hours
Estimated Duration: 7 weeks
```

Project Plan Link

```
MO-IT101 | Worksheet - | Group 9_Project Plan
```

---

# Project Plan Note

The project plan served as a guideline for organizing the development tasks and timeline of the MotorPH Payroll System. However, the schedule was not followed strictly according to the exact dates indicated in the document.

Since team members were working on multiple subjects simultaneously, some development activities were completed earlier or later than planned.

Additionally, while development work continued throughout the project, some updates and progress were not immediately reflected in the project plan sheet. As a result, certain task statuses in the document may not fully represent the actual progress of the implementation.

Despite these adjustments, the project plan still helped structure the workflow and identify the major components required for the development of the payroll system.

---

# Testing and Revision

During development, testing and revision were primarily handled by **Kimberly Joy Goyena**, while **Harty Joy Villegas** worked on preparing and finalizing the project documentation and README.

Testing activities included:

- Performing unit testing and debugging techniques for each payroll component
- Conducting integration testing of the full payroll process
- Validating edge cases such as zero hours worked and high overtime scenarios
- Fixing identified bugs and logic errors
- Refactoring code for readability and efficiency
- Improving input validation and error handling

These activities helped ensure the correctness and stability of the payroll computation before submission.

---

# Current Status

This repository represents the **Milestone 2 Initial Code Implementation**.

Currently implemented features:

- Excel file reading
- Attendance record processing
- Salary computation
- Government deduction simulation
- Console-based payroll output

Possible future improvements include:

- improved system modularization
- input validation improvements
- graphical user interface implementation
- database integration

---

# Milestone 1 Alignment and Mentor Feedback

This project builds upon the system planning and documentation completed during **Milestone 1**.

The initial submission included:

- Team Details
- MotorPH Requirements
- Use Case Diagram
- Wireframe
- Effort Estimation
- Project Plan

These documents served as the foundation for the system implemented in **Milestone 2**.

---

## Mentor Feedback

Instructor: **Aldrin John Tamayo**

| Section | Feedback |
|------|------|
| Team Details | All members were listed clearly and the group number was identified. |
| MotorPH Requirements | Requirements were correct but descriptions were too broad for Phase 1 implementation. |
| Use Case Diagram | Correct actors and use cases were identified (Employee and Payroll Staff). |
| Wireframe | Correctly presented as a functional sketch without design elements. |
| Effort Estimation | Time estimates were realistic and development challenges were identified. |
| Project Plan | Clear weekly breakdown of tasks and responsibilities. |

---

# Requirement Improvement Based on Feedback

Original requirement descriptions were broad:

- Employee Information Management
- Salary Computation System

Based on the mentor’s recommendation, the requirements were revised to focus on a simpler **Phase 1 implementation**.

Revised requirements:

### Presentation of Employee Details

- Employee ID
- Employee Name
- Employee Birthday

### Automatic Salary Calculation

- Total Hours Worked
- Gross Salary
- Net Salary

These revised requirements guided the development of the **Milestone 2 payroll system implementation** contained in this repository.
