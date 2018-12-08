# EverydaycardYNABConverter
Powershell script that converts an Everydaycard bank Excel file so it can be imported by YNAB.

### Usage
```@echo off
Powershell.exe -executionpolicy remotesigned -File ConvertXlsToCsv.ps1 "kontotransactionlist.xls" "output_everydaycard.csv"
pause
```

First argument specified the input, second argument specified the output.

### Requirements
ACE Driver (Download and install `Microsoft Access Database Engine 2010 Redistributable`)
- Required to convert XLS file to CSV without having Microsoft Excel installed
