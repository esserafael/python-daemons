<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.145
	 Created on:   	24/08/2020 11:31
	 Created by:   	07574534900 - Rafael Feustel
	 Organization: 	Uniasselvi
	 Filename:     	ConvertTo-ExcelCustomReport.ps1
	===========================================================================
	.DESCRIPTION
		Creates an Excel spreadsheet using a custom CSV report to send daily by e-mail.
#>

### Set input and output path
$InputHTML = "C:\Users\07574534900\Dropbox\1 - Uniasselvi\Scripts\GitHub\python-daemons\ms-graph-api-UserActivity\teste.html"
$OutputXLSX = "C:\Users\07574534900\Dropbox\1 - Uniasselvi\Scripts\GitHub\python-daemons\ms-graph-api-UserActivity\teste2.xlsx"

$Excel = New-Object -ComObject Excel.Application
$Excel.Workbooks.Open($InputHTML)
#($Excel.Workbooks[1]).Activate()
$Excel.ActiveWorkbook.SaveAs($OutputXLSX, 51)
$Excel.ActiveWorkbook.Close()
$Excel.Quit()