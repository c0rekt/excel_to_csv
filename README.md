# excel_to_csv
A PowerShell function that converts excel to csv

Requirements

-PowerShell

-Activated Microsoft Excel(Unactivated Microsoft does not work)

Purpose

The purpose of this program is to convert Excel extension files to CSV format.

Syntax

-ExcelFileName

--Excel file name together with its extension should be entered.

-ExcelFilePath

--File path of the excel file that has been declared in –ExcelFileName.

-SheetName

--The sheet name that is in the workbook (Note: Even if there is only 1 sheet, the sheet name must still be added in).

-CsvLocation

--The directory the csv should be saved at.

-Delimiter

--The character to separate your values.

-RowStart

--The number where the workbook row actually starts

  

How to use

Firstly, import the module.

Secondly, input the statement with the right syntaxes. An example is shown below:

ExcelToCsv -ExcelFileName "example.xlsx" -Delimiter "|" -CsvLocation "C:\Users\test\Desktop\test\" -RowStart 3 -ExcelFilePath "C:\Users\test\Desktop\test\" -SheetName "Sheet1"

