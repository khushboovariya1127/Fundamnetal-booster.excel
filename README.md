# 📊 Fundamental Booster — Excel Project

## 📌 Project Overview

**Fundamental Booster** is an Excel practice project designed to strengthen fundamental and intermediate Excel skills through practical data-analysis tasks.

The workbook contains multiple datasets and formula examples covering:

- Student grade analysis
- Sales analysis
- Employee analysis
- Excel lookup and reference functions
- Conditional formulas
- Date calculations
- Aggregation and criteria-based calculations
- Dynamic and reference functions

This project is useful for learning how Excel formulas can be applied to real-world-style datasets.

---

## 🎯 Project Goal

The main goal of this project is to practice and understand commonly used Excel functions by applying them to structured datasets.

The workbook demonstrates how to:

- Classify data using conditions
- Calculate grades and eligibility
- Analyze dates and years of service
- Calculate discounts and net sales
- Search and retrieve information from tables
- Perform conditional calculations
- Work with text, date, mathematical, lookup, and reference functions

---

## 📁 Workbook Structure

| Sheet | Description |
|---|---|
| **Students** | Student marks, grades, eligibility, and enrollment analysis |
| **Sales** | Sales amount, discounts, net sales, and product-price lookup |
| **Employees** | Employee salary, department, joining-date, and service analysis |
| **Functions** | Examples and practice reference for important Excel functions |

---

## 👨‍🎓 Students Sheet

The **Students** sheet contains student academic information including:

- Student ID
- Student Name
- Mathematics marks
- Science marks
- English marks
- Enrollment Date
- Grade
- Math eligibility
- Math & Science eligibility
- Years since enrollment

### Functions Practiced

- `IF`
- `AVERAGE`
- `AND`
- `COUNTIFS`
- `FILTER`
- `AVERAGEIFS`
- `DATEDIF`
- Date calculations

### Example Tasks

- Assign grades based on average marks.
- Check whether a student's Math score is above 60.
- Check whether Math and Science scores satisfy eligibility conditions.
- Calculate the number of years since enrollment.
- Count students based on multiple criteria.
- Filter students according to conditions.
- Calculate averages using multiple criteria.

### Grade Classification

| Average Score | Grade |
|---:|:---|
| 80 or above | A |
| 60–79 | B |
| 50–59 | C |
| Below 50 | D |

---

## 💰 Sales Sheet

The **Sales** sheet is used to practice sales calculations and lookup functions.

It contains information such as:

- Sales ID
- Product
- Region
- Salesperson
- Sales Amount
- Sales Date
- Discount Percentage
- Discount Amount
- Net Sales
- Product Price reference table

### Functions Practiced

- `IF`
- `VLOOKUP`
- `XLOOKUP`
- `SUMIFS`
- `AVERAGEIFS`
- `INDEX`
- `MATCH`
- `XMATCH`
- `INDIRECT`
- `OFFSET`

### Discount Logic

| Sales Amount | Discount |
|---:|---:|
| ₹40,000 or above | 10% |
| ₹20,000–₹39,999 | 5% |
| Below ₹20,000 | 0% |

### Sales Calculations

**Discount Amount:**

```excel
=Amount*Discount%
```

**Net Sales:**

```excel
=Amount-Discount Amount
```

The sheet also demonstrates how product prices can be retrieved from a reference table using lookup functions.

---

## 👩‍💼 Employees Sheet

The **Employees** sheet contains employee-related information including:

- Employee ID
- Employee Name
- Department
- Salary
- Joining Date
- Years of Service
- Salary Band

It also contains a lookup section for retrieving employee information.

### Functions Practiced

- `XLOOKUP`
- `DATEDIF`
- `IF`
- Lookup and reference functions

### Salary Band Logic

| Salary | Salary Band |
|---:|:---|
| ₹100,000 or above | High |
| ₹60,000–₹99,999 | Medium |
| Below ₹60,000 | Low |

### Employee Analysis

The project calculates:

- Years of service from joining date
- Salary category
- Employee salary using Employee ID
- Employee name using Employee ID
- Employee department using Employee ID

---

## 🧮 Functions Sheet

The **Functions** sheet provides a reference for commonly used Excel functions.

### Conditional Functions

- `IF`
- `AND`
- `OR`

### Criteria-Based Functions

- `COUNTIFS`
- `SUMIFS`
- `AVERAGEIFS`

### Lookup Functions

- `VLOOKUP`
- `XLOOKUP`
- `INDEX`
- `MATCH`
- `XMATCH`

### Date Functions

- `DATEDIF`
- Date calculations

### Reference Functions

- `INDIRECT`
- `OFFSET`

### Text Functions

The workbook also provides examples for working with text and string-related operations.

---

## 🛠️ Excel Skills Demonstrated

This project demonstrates practical knowledge of:

- Conditional formatting and classification
- Formula-based data analysis
- Multiple-condition logic
- Lookup and reference operations
- Date calculations
- Sales calculations
- Salary analysis
- Student performance analysis
- Criteria-based aggregation
- Dynamic references
- Data filtering

---

## 📚 Learning Outcomes

After completing this project, the learner should be able to:

1. Use `IF` to create conditional results.
2. Combine conditions using `AND` and `OR`.
3. Calculate averages and totals.
4. Use `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
5. Retrieve values using `XLOOKUP` and `VLOOKUP`.
6. Understand `INDEX`, `MATCH`, and `XMATCH`.
7. Calculate dates and years of service using `DATEDIF`.
8. Create sales discount and net-sales calculations.
9. Categorize employees according to salary.
10. Analyze student performance using formulas.
11. Use Excel reference functions such as `INDIRECT` and `OFFSET`.
12. Build a foundation for more advanced Excel data analysis.

---

## 📂 Project File

The main workbook for this project is:

**`Fundamental_Booster.xlsx`**

---

## 🚀 How to Use

1. Open `Fundamental_Booster.xlsx` in Microsoft Excel.
2. Review the datasets in the **Students**, **Sales**, and **Employees** sheets.
3. Examine the formulas used in each calculated column.
4. Open the **Functions** sheet to review the Excel functions.
5. Modify sample data and observe how the formulas update.
6. Use the workbook as a practice project for Excel and data-analysis skills.

---

## 🧠 Skills & Technologies

**Tool:** Microsoft Excel

**Topics:**

- Excel Formulas
- Data Analysis
- Conditional Logic
- Lookup Functions
- Date Functions
- Reference Functions
- Criteria-Based Calculations
- Spreadsheet Data Management

---

## 👤 Project Type

**Excel Fundamental & Data Analysis Practice Project**

This project is intended for learning, practice, portfolio development, and preparation for entry-level data-analysis tasks.

---

## ⭐ Conclusion

This project provides hands-on practice with important Excel functions and data-analysis techniques. It demonstrates how formulas can be used to transform raw data into meaningful information and build a strong foundation for advanced Excel and data-analysis projects.

⭐ **If you find this project useful, feel free to explore and build upon it with advanced Excel dashboards and data-analysis projects.**
