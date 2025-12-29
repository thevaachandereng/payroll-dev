# Malaysia Payroll + Payslip Generator (Shiny)

A complete **Malaysia payroll calculator + payslip generator** built in **R Shiny**, implementing:

- Full **SOCSO Act 4 (First Category)** table
- Full **EIS Act 800** table
- **EPF** (employee + employer, configurable rates)
- **PCB income tax (annual-bracket based approximation)** with marital status & children
- Attendance allowance, OT, Sunday pay logic
- Excel-based payslip generation

This README documents the **current logic exactly as implemented in the latest code**.

---

## 1. Features Overview

- Upload Excel payroll inputs or use built-in sample data
- Configurable company rules (OT, Sunday, attendance allowance, EPF rates)
- Correct separation of:
  - **Statutory wage** (for SOCSO / EIS / PCB)
  - **Gross salary** (display)
  - **Final pay** (what employee receives)
- Batch export:
  - Payroll table (CSV / Excel)
  - Individual payslips (Excel, zipped)

---

## 2. Required Input Columns (Excel)

| Column | Description |
|------|-------------|
| STAFF NAME | Employee name |
| ACCOUNT NUMBER | Bank account number |
| IDENTIFICATION CARD | IC number |
| BASIC | Monthly basic salary (RM) |
| ANNUAL LEAVE | Informational only |
| ABSENCE | Number of unpaid leave days |
| OT | Overtime hours |
| SUNDAY | Sunday hours |
| ALLOWANCE TRANSPORT | Transport allowance (RM) |
| CASH ADVANCE COMPANY | Salary advance deduction |
| CASH ADVANCE MANAGER | Salary advance deduction |
| MARITAL STATUS | single / married_spouse_working / married_spouse_not_working |
| CHILDREN | Number of children |

Missing `ACCOUNT NUMBER` or `IDENTIFICATION CARD` are auto-filled as blank.

---

## 3. Working Time & Allowance Rules

### Working time
- Working days: configurable (default **26**)
- Hours per day: configurable (default **8**)

### Attendance allowance
Attendance allowance is **paid only if**:
- `ABSENCE == 0`
- `BASIC <= Attendance Threshold`

Otherwise, attendance allowance = **0**.

---

## 4. Overtime (OT) Calculation

### OT base salary
- **Basic only** → `BASIC`
- **Basic + Attendance Allowance** → only if `ABSENCE == 0`

```text
OT/HOUR = (OT Base Salary / Working Days / Hours Per Day) × OT Multiplier
OVERTIME PAY = OT/HOUR × OT Hours
```

Attendance allowance **can be included for OT** if selected.

---

## 5. Sunday Pay Calculation

Sunday pay is **always calculated on BASIC only**:

```text
SUNDAY PAY = (BASIC / Working Days / Hours Per Day) × Sunday Multiplier × Sunday Hours
```

Attendance allowance is **never included** in Sunday pay.

---

## 6. Statutory Wage vs Gross Salary

### Statutory Wage (used for SOCSO / EIS / PCB)

```text
STATUTORY WAGE =
  (BASIC − NO PAY LEAVE)
  + OVERTIME
  + SUNDAY PAY
```

Attendance allowance is **excluded** by design.

### Gross Salary (display only)

```text
GROSS SALARY =
  STATUTORY WAGE
  + ATTENDANCE ALLOWANCE
  + ALLOWANCE TRANSPORT
```

---

## 7. EPF Calculation

EPF is calculated **on BASIC − NO PAY LEAVE only**:

```text
EMPLOYEE EPF = (BASIC − NO PAY LEAVE) × Employee EPF Rate
EMPLOYER EPF = (BASIC − NO PAY LEAVE) × Employer EPF Rate
```

Default rates:
- Employee: **11%**
- Employer: **13%**

---

## 8. SOCSO (Act 4 – First Category)

- Full **Act 4 table** implemented
- Lookup is done using:

```text
wage > wage_low AND wage <= wage_high
```

SOCSO is calculated on:
```text
STATUTORY WAGE
```

Outputs:
- Employer SOCSO
- Employee SOCSO

---

## 9. EIS (Act 800)

- Full **Act 800 table** implemented
- Same lookup logic as SOCSO

EIS is calculated on:
```text
STATUTORY WAGE
```

Outputs:
- Employer EIS
- Employee EIS

---

## 10. PCB Income Tax Logic (Important)

### Annual-based PCB approximation

PCB is calculated as:

1. Convert **STATUTORY WAGE** to annual income
2. Subtract reliefs:
   - Self relief: RM 9,000
   - Spouse relief: RM 4,000 (only if spouse not working)
   - Child relief: RM 2,000 × children
   - EPF relief: capped at RM 4,000 annually
3. Apply **official tax brackets**
4. Apply RM 400 rebate if chargeable income ≤ RM 35,000
5. Divide annual tax by 12

```text
PCB (monthly) = (Annual Tax − Rebate) / 12
```

### ⚠️ PCB Disclaimer

This is a **simplified annual-bracket PCB model**.

Differences vs LHDN / Salary.gov may occur because:
- Real PCB tables do **not always apply child relief** monthly
- LHDN uses **PCB tables**, not pure marginal tax formulas

This model is suitable for **internal payroll estimation**, not statutory filing.

---

## 11. Net Pay & Final Pay

### Nett Pay (after statutory deductions)

```text
NETT PAY =
  STATUTORY WAGE
  − EMPLOYEE EPF
  − EMPLOYEE SOCSO
  − EMPLOYEE EIS
  − INCOME TAX
```

### Final Pay (employee receives)

```text
FINAL PAY =
  NETT PAY
  + ATTENDANCE ALLOWANCE
  + ALLOWANCE TRANSPORT
  − CASH ADVANCE COMPANY
  − CASH ADVANCE MANAGER
```

---

## 12. Total Employer Cost

```text
TOTAL EMPLOYER COST =
  STATUTORY WAGE
  + ATTENDANCE ALLOWANCE
  + ALLOWANCE TRANSPORT
  + EMPLOYER EPF
  + EMPLOYER SOCSO
  + EMPLOYER EIS
```

---

## 13. Payslip Generation

- Uses `employeeslip.xlsx` as a template
- Writes values to fixed cells
- Generates one Excel payslip per employee
- Zips all payslips for download

### Template requirements
- Must contain a worksheet named **Template**
- Cell positions must match the code

---
