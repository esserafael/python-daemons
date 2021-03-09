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
	Add-Content -Path $LogPath -Value "$($DateTimeNow) - $($Text)"
}

$ScriptPath = Split-Path ($MyInvocation.MyCommand.Path)
$LogPath = "$($ScriptPath)\debug_PS.log"

Write-LocalLog -Text "Script started."

$MarvinCredPath = "$($ScriptPath)\marvin@uniasselvi.com.br.xml"

try
{
	$MarvinCred = Import-Clixml -Path $MarvinCredPath
	Write-LocalLog "Office 365 SMTP Cred loaded from encrypted data in file '$($MarvinCredPath)'."
}
catch [System.IO.FileNotFoundException]
{
	Write-LocalLog "$($_.Exception.Message)"
}
catch [System.Security.Cryptography.CryptographicException]
{
	Write-LocalLog "$($_.Exception.Message)"
}
catch
{
	# The catches do nothing different anyways...
	Write-LocalLog "$($_.Exception.Message)"
}

$XlsxFileCreated = $false

try
{
	$Excel = New-Object -ComObject Excel.Application
	Write-LocalLog -Text "Excel COM Object created."
	
	Write-LocalLog -Text "Opening html file: '$($HtmlPath)'."
	$Excel.Workbooks.Open($HtmlPath)
	Write-LocalLog -Text "Html file opened."
	
	Write-LocalLog -Text "Saving Xlsx file: '$($XlsxPath)'."
	$Excel.ActiveWorkbook.SaveAs($XlsxPath, 51)
	
	$XlsxFileCreated = $true	
	Write-LocalLog -Text "Xlsx file saved."
	
	$Excel.ActiveWorkbook.Close($false)
	$Excel.Quit()
	
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
	
	Write-LocalLog -Text "Conversion completed."
}
catch
{
	Write-LocalLog -Text "Error in conversion: $($_.Exception.Message)"
}

if ($XlsxFileCreated)
{
	#$To = @("paula.rodrigues@uniasselvi.com.br", "pedro.graca@uniasselvi.com.br", "cloves.machado@uniasselvi.com.br")
	$To = @("rafael.gustmann@uniasselvi.com.br")
	$DateString = Split-Path $XlsxPath -Leaf | Select-String -Pattern "^.+_(\d+-\d+-\d+)_"
	
	try
	{
		Send-MailMessage `
						 -From $MarvinCred.UserName `
						 -To $To `
						 -Cc "rafael.gustmann@uniasselvi.com.br" `
						 -Subject "Relatório diário de acessos ao Office 365." `
						 -Body "<p style='font-family: 'Segoe UI';'>Bom dia!<br /><br />Em anexo está o arquivo com o relatório diário de acessos ao Office 365, do dia $($DateString.Matches.Groups[1].Value).</p>" `
						 -BodyAsHtml `
						 -Encoding UTF8 `
						 -Attachments $XlsxPath `
						 -SmtpServer "smtp.office365.com" `
						 -Port "587" `
						 -UseSsl `
						 -Credential $MarvinCred `
						 -ErrorAction Stop
		
		Write-LocalLog -Text "Mail sent to '$($To)'."
	}
	catch
	{
		Write-LocalLog -Text "Error sending mail: $($_.Exception.Message)"
	}

	try
	{
		$Compress = @{
			Path = $XlsxPath, $HtmlPath, ($HtmlPath -replace ".html", ".csv")
			CompressionLevel = "Optimal"
			DestinationPath = ($XlsxPath -replace ".xlsx", ".zip")
		}
	
		Compress-Archive @Compress -ErrorAction Stop
		
		try
		{
			Remove-Item -Path @($XlsxPath, $HtmlPath, ($HtmlPath -replace ".html", ".csv")) -ErrorAction Stop
		}
			catch
		{
			Write-LocalLog -Text "Error removing files after compression: $($_.Exception.Message)"
		}		
	}
	catch
	{
		Write-LocalLog -Text "Error compressing files: $($_.Exception.Message)"
	}
}
