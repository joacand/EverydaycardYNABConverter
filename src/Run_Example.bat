@echo off
Powershell.exe -executionpolicy remotesigned -File ConvertXlsToYnabCsv.ps1 "kontotransactionlist.xls" "output_everydaycard.csv"
pause