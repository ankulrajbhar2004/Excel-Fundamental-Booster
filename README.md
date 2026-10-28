# Project Fundamental Booster

An Excel practice workbook designed to strengthen fundamental
spreadsheet skills through student performance analysis, sales
calculations, and employee salary lookups.

## Overview

**Project Fundamental Booster** contains three worksheets with sample
datasets and formula-based exercises:

-   **Student Data** --- calculate total marks, assign grades, check
    subject-score conditions, calculate percentages, and practice
    rounding functions.
-   **Sales Data** --- calculate discounts, determine eligibility and
    performance labels, summarize sales by product and region, and look
    up product details.
-   **Employee Data** --- explore employee information and retrieve
    salaries by employee name.

## Workbook Contents

### 1. Student Data

The student worksheet includes Student ID, Name, Math, Science, English,
Enrollment Date, and calculated columns.

Excel concepts practiced: - `SUM` / arithmetic formulas for total
marks - `IF` and nested `IF` for grade assignment - `AND` for checking
multiple subject-score conditions - `ROUND` for percentage formatting -
`CEILING` and `FLOOR` for rounding marks to a chosen multiple -
Filtering students by percentage or total score

### 2. Sales Data

The sales worksheet includes Sales ID, Product, Region, Salesperson,
Amount, Date, Discount, Discount Amount, Eligibility, and Performance.

Excel concepts practiced: - `IF` and nested `IF` for discount tiers and
performance labels - `OR` for eligibility conditions - `SUMIFS` to total
sales for a specific product and region - `ROUND`, `CEILING`, and
`FLOOR` - `VLOOKUP` for retrieving product information from a sales ID -
`INDEX` and `MATCH` for multi-criteria lookups, such as salesperson and
month

### 3. Employee Data

The employee worksheet includes Employee ID, Name, Department, Salary,
and Joining Date.

Excel concepts practiced: - `XLOOKUP` to retrieve an employee's salary
by name - Working with structured table references - Searching records
without manually sorting the data

> Note: `XLOOKUP` requires a recent version of Microsoft Excel. In older
> Excel versions, an `INDEX` + `MATCH` formula can be used as an
> alternative.

## Skills Demonstrated

-   Excel tables and structured references
-   Logical functions: `IF`, `AND`, `OR`
-   Conditional aggregation: `SUMIFS`, `AVERAGEIFS`
-   Lookup functions: `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`
-   Mathematical and rounding functions: `ROUND`, `CEILING`, `FLOOR`
-   Data filtering and basic analysis
-   Formula troubleshooting and error checking

## How to Use

1.  Download or clone this repository.
2.  Open `Project Fundamental Booster.xlsx` in Microsoft Excel.
3.  Review each worksheet and its sample data.
4.  Select a formula cell to inspect the formula in the formula bar.
5.  Change input values or criteria to practice how results update.
6.  Try extending the tables with additional records and copy formulas
    into new rows.

## Requirements

-   Microsoft Excel is recommended for full compatibility with Excel
    tables and structured references.
-   Excel for Microsoft 365 or a recent Excel version is recommended for
    `XLOOKUP` and dynamic-array functions such as `FILTER`.

Some formulas may behave differently in other spreadsheet applications.
Verify formulas after opening the workbook in Google Sheets or
LibreOffice Calc.

## Repository Structure

``` text
Project-Fundamental-Booster/
├── Project Fundamental Booster.xlsx
└── README.md
```

## Learning Goal

The goal of this workbook is to build confidence with commonly used
Excel formulas and apply them to realistic student, sales, and employee
datasets.

## Author

**Ankul Rajbhar**

------------------------------------------------------------------------

