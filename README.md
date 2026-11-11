# Online Retail Excel Project

## Files
- Online_Retail_Project.xlsx : main Excel workbook
- PowerQuery_Sales_Query.txt : ready Power Query M code
- RefreshAndFormat.bas : VBA macro code
- README.md : instructions

## Kaggle Dataset
Download Online Retail.xlsx from:
https://www.kaggle.com/datasets/samantas2020/online-retail-xlsx

Keep the downloaded file in the same folder as this project.

## What This Project Does
- Imports Online Retail.xlsx using Power Query
- Cleans invalid quantities and prices
- Removes blank and duplicate records
- Creates Sales Amount column
- Separates cancellations/returns using InvoiceNo starting with C
- Merges sales and returns data
- Loads final data to Cleaned_Data sheet
- Creates Pivot Table for Country and Product analysis
- Uses VBA macro to refresh and format data
- Saves final file as .xlsm

## Resume Line
Built an Excel-based Online Retail analysis project using Power Query for data cleaning, merging, and transformation; Pivot Tables for country and product-level sales analysis; and VBA for automated refresh and formatting.
