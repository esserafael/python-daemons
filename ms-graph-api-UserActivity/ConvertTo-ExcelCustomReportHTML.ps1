<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.145
	 Created on:   	24/08/2020 11:31
	 Created by:   	07574534900 - Rafael Feustel
	 Organization: 	Uniasselvi
	 Filename:     	ConvertTo-ExcelCustomReportHTML.ps1
	===========================================================================
	.DESCRIPTION
		Creates an Excel file (XLSX) using a custom Html file/report and send daily by e-mail.
#>

Param (
	[Parameter(Mandatory = $true)]
	[String]$HtmlPath,
	[Parameter(Mandatory = $true)]
	[String]$XlsxPath
)

function Write-LocalLog
{
	Param (
		[String]$Text
	)
	
	$DateTimeNow = Get-Date -f "yyyy/MM/dd HH:mm:ss.fff"
	Add-Content -Path $LogPath -Value "$($DateTimeNow) - $($Text)."
}

$ScriptPath = Split-Path ($MyInvocation.MyCommand.Path)
$LogPath = "$($ScriptPath)\debug_PS.log"

Write-LocalLog -Text "Script started."

try
{
	$Excel = New-Object -ComObject Excel.Application
	Write-LocalLog -Text "Excel COM Object created."
	
	$Excel.Workbooks.Open($HtmlPath)
	Write-LocalLog -Text "Html file opened: $($HtmlPath)"
	
	#($Excel.Workbooks[1]).Activate()
	$Excel.ActiveWorkbook.SaveAs($XlsxPath, 51)
	Write-LocalLog -Text "Xlsx file saved: '$($XlsxPath)'."
	
	$Excel.ActiveWorkbook.Close($false)
	$Excel.Quit()
	
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
	
	Write-LocalLog -Text "Conversion completed."
}
catch
{
	Write-LocalLog -Text "Error in conversion: $($_.Exception.Message)"
}
