@echo off
Powershell.exe -executionpolicy remotesigned -File ConvertXlsToYnabCsv.ps1 "Transaktionslista.xls" "output_everydaycard.csv"
pause
