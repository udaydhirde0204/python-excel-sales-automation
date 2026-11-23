# Automated Sales Report Generator (Python + Excel)

A Python automation tool built with **Pandas** and **OpenPyXL** that converts raw, uncleaned Excel sales records into a formatted, multi-sheet executive dashboard with automated charts.

## Business Problem Solved
Small businesses and retail teams spend hours every week manually cleaning spreadsheets, fixing typos, calculating totals, and building charts. This script automates the entire workflow in under 2 seconds.

## Key Features
- **Automated Data Cleaning:** Removes invalid dates, negative quantities, and standardizes inconsistent product/category formatting.
- **KPI & Revenue Calculations:** Automatically computes total revenue, order counts, units sold, and top-performing products.
- **Multi-Sheet Excel Output:** Generates `Dashboard`, `Monthly Summary`, `Product Summary`, and `Cleaned Data` tabs.
- **Visual Charts:** Embeds automated Monthly Revenue and Product Revenue charts directly onto the Dashboard.

## Tech Stack
- **Python 3.12**
- **Pandas** (data cleaning & aggregation)
- **OpenPyXL** (Excel styling, currency formatting & charts)

## How to Run
1. Install dependencies:
   ```bash
   pip install pandas openpyxl
