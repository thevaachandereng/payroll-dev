# 🇲🇾 Malaysia Payroll System (Shiny + Excel Payslips)

A fully parameterized **Malaysia payroll system** built in **R + Shiny**, supporting statutory deductions (EPF, SOCSO, EIS), PCB income tax estimation, and automated **Excel payslip generation** using a fixed Excel template.

This repository contains a production-ready payroll engine suitable for internal HR use, auditing, and further extension.

---

## 📌 Key Features

- 📥 Upload payroll data via Excel (.xlsx)
- 🧮 Automated payroll calculations
  - Basic salary adjustments
  - Overtime (OT)
  - Sunday pay (Field vs Office logic)
  - Attendance allowance logic
- 🏛 Statutory deductions
  - EPF (Employee & Employer)
  - SOCSO (Act 4 – First Category)
  - EIS (Act 800)
- 💰 Monthly PCB income tax (approximation)
- 📊 Outputs
  - Payroll preview table (Shiny UI)
  - CSV export
  - Excel export
  - Individual employee payslips (Excel)
- 📦 Batch payslip download as ZIP

---

## 📂 Project Structure

```
project/
│
├── app.R                     # Main Shiny application
├── employeeslip.xlsx         # Payslip Excel template (required)
├── README.md                 # Documentation (this file)
└── sample_payroll.xlsx       # Example input file
```

---

## 🧾 Input Data Requirements

Uploaded Excel files must contain the following columns.  
Missing columns will be **automatically created and filled with defaults (0 or blank)**.

### 🔑 Employee Identifiers

| Column | Description |
|------|------------|
| STAFF NAME | Employee full name |
| IDENTIFICATION CARD | NRIC / Passport |
| ACCOUNT NUMBER | Bank account |
| STAFF TYPE | Office / Field |
| MARITAL STATUS | single / married_spouse_not_working |
| CHILDREN | Number of dependents |

### 💵 Salary & Work Data

| Column | Description |
|------|------------|
| BASIC | Monthly basic salary |
| ANNUAL LEAVE | Informational |
| ABSENCE | Unpaid leave days |
| MEDICAL LEAVE | Sick leave days |
| OT | Overtime hours |
| SUNDAY | Sunday days/hours worked |

### ➕ Allowances

| Column |
|------|
| ALLOWANCE TRANSPORT |
| ALLOWANCE |
| CASH ALLOWANCE |
| BIKE ALLOWANCE |

### ➖ Deductions / Advances

| Column |
|------|
| CASH ADVANCE COMPANY |
| CASH ADVANCE MANAGER |
| ADVANCE-DIRECT PBB TRANSFER |
| ALLOWANCE-PREPAYMENT |

---

## ⚙️ Payroll Calculation Logic

### 1️⃣ Daily & Hourly Pay

```
Daily Pay    = BASIC / Working Days
Hourly Rate = Daily Pay / Hours per Day
```

### 2️⃣ Unpaid Leave

```
NO PAY LEAVE   = Daily Pay × ABSENCE
BASIC PAYABLE = BASIC − NO PAY LEAVE
```

### 3️⃣ Overtime (OT)

**Office Staff**
```
OT = OT Hours × Hourly Rate × OT Multiplier
```

### 4️⃣ Sunday Pay

**Office**
```
Sunday Pay = Hourly Rate × Sunday Multiplier × Hours
```

**Field**
```
Sunday Pay = (Basic Payable / Working Days) × Sunday Multiplier × Days
```

### 5️⃣ Attendance Allowance

Attendance allowance is forfeited if:
- Any unpaid absence
- Any medical leave
- Basic salary exceeds threshold

```
Attendance Allowance =
  if (ABSENCE > 0 OR MEDICAL LEAVE > 0 OR BASIC > Threshold)
    0
  else
    Allowance Amount
```

---

## 🏛 Statutory Contributions

### EPF
- Calculated on **Basic Payable**
- Rates configurable in UI
- Rounded up using `ceiling()`

### SOCSO (Act 4)
- Wage-band based
- Applied to **Statutory Wage**
- Employee & employer portions calculated separately

### EIS (Act 800)
- Wage-band based
- Applied to **Statutory Wage**

---

## 💰 Income Tax (PCB – Approximation)

PCB is estimated using:
- Annualized statutory wage
- EPF relief (capped)
- Spouse & child reliefs

⚠️ This is an **estimate only**, not an LHDN-certified PCB calculator.

---

## 🧾 Salary Aggregates

### Gross Salary

```
GROSS SALARY =
  STATUTORY WAGE
+ TOTAL ALLOWANCES
```

### Total Deductions

```
TOTAL DEDUCTIONS =
  EPF (Employee)
+ SOCSO (Employee)
+ EIS (Employee)
+ Income Tax
+ Advances / Prepayments
```

### Final Pay

```
NETT PAY =
  STATUTORY WAGE
− EPF
− SOCSO
− EIS
− Income Tax

FINAL PAY =
  NETT PAY
+ Allowances
− Advances
```

---

## 📄 Payslip Excel Template Mapping

The app writes values directly into fixed Excel cells in `employeeslip.xlsx`.

### Earnings (Left)

| Cell | Value |
|----|------|
| B6 | Basic Payable |
| B7 | Sunday Pay |
| B8 | Transport Allowance |
| B9 | Attendance Allowance |
| B10 | OT + Other Allowances |
| **B12** | **Gross Salary (TOTAL)** |

### Deductions (Right)

| Cell | Value |
|----|------|
| D6 | EPF |
| D7 | SOCSO |
| D8 | Advances |
| D9 | Allowance Prepayment |
| D10 | Unpaid Leave |
| D11 | EIS |
| **D12** | **Total Deductions** |

### Final Pay

| Cell | Value |
|----|------|
| D18 | Final Net Pay |

⚠️ **Do not modify the Excel layout** unless cell mappings are updated in code.

---

## ▶️ Running the Application

```r
shiny::runApp("app.R")
```

Ensure `employeeslip.xlsx` is located in the same directory as `app.R`.

---

## ⚠️ Disclaimer

This system is intended for **internal payroll estimation only**.