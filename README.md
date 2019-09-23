# Excel-Automated-Invoice
Excel VBA Automated Invoice Template

## Overview
Uses Visual Basics and Excel to automatically copy invoice information from csv (Comma-separated values) format to a customized template.

## Instructions
1)  Copy csv information into Sheet1 under header.
2)  Run Script Clean_Data_Up if necessary to format csv data before invoice creation.
3)  Run Script Automated_Invoice to automatically copy formatted csv data into Sheet2 invoice templet.

## Version Update and Fixes
V.02  
Added remove carrier USPS from Service Selected.  
Added removal of previous data from invoice template.  
Added new Array for shipping service price addition base on service selected.  
Fixed Array for combining City, State, Zip. Array read from only row 2.  
Fixed For() Loop would run until set number. Now For() Loop runs until last row in column A that has data.  
Fixed City, State, Zip text would insert into Sheet1 if Sheet2 wasn't active sheet.  
