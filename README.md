# Excel-Automated-Invoice
Excel VBA Automated Invoice Template

## Overview
Uses Visual Basics and Excel to automatically copy invoice information from csv (Comma-separated values) format to a customized template.

## Instructions
1)  Copy csv information into Sheet1 under header.
2)  Run Script Clean_Data_Up if necessary to format csv data before invoice creation.
3)  Run Script Automated_Invoice to automatically copy formatted csv data into Sheet2 invoice templet.
4)  Run Script Invoice_Tracker to scan and highlight any issues with invoices.

## Version Update and Fixes
V.03
Added Invoice_Tracker - New module to track invoices created
Added Automated_Invoice - Copy Invoice Number, Date, and Item Qty To Invoice Tracker
Fixed Clean_Data_Up - Remove PO from Order Number was not bound by Range constraints.
Removed Automated_Invoice - Print invoice to printer and "Click To Print" message box. I found this redundent and slow because you can print from PDF files afterwards.

V.02  
Added Clean_Data_Up - remove carrier USPS from Service Selected.  
Added Automated_Invoice - removal of previous data from invoice template.  
Added Automated_Invoice - new Array for shipping service price addition base on service selected.  
Fixed Automated_Invoice - Array for combining City, State, Zip. Array read from only row 2.  
Fixed Automated_Invoice - For() Loop would run until set number. Now For() Loop runs until last row in column A that has data.  
Fixed Automated_Invoice - City, State, Zip text would insert into Sheet1 if Sheet2 wasn't active sheet.
