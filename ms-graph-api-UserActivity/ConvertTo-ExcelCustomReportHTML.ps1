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

function Compress-ReportFiles
{
	Param (
		[System.Collections.Hashtable]$Compress
	)

	try
	{
		Compress-Archive @Compress -ErrorAction Stop	
		Write-LocalLog -Text "Files compressed."
	}
	catch
	{
		Write-LocalLog -Text "Error compressing files: $($_.Exception.Message)"
	}
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

	$Excel.Visible = $true
	
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
	$Compress = @{
		Path = $XlsxPath
		CompressionLevel = "Optimal"
		DestinationPath = ($XlsxPath -replace ".xlsx", ".zip")
	}

	Compress-ReportFiles -Compress @Compress

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
						 -Attachments ($XlsxPath -replace ".xlsx", ".zip") `
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

	$Compress = @{
		Path = $HtmlPath, ($HtmlPath -replace ".html", ".csv")
		CompressionLevel = "Optimal"
		DestinationPath = ($XlsxPath -replace ".xlsx", "_sources.zip")
	}

	Compress-ReportFiles -Compress @Compress
	
	try
	{
		Remove-Item -Path @($XlsxPath, $HtmlPath, ($HtmlPath -replace ".html", ".csv")) -ErrorAction Stop
	}
		catch
	{
		Write-LocalLog -Text "Error removing files after compression: $($_.Exception.Message)"
	}		
}

# SIG # Begin signature block
# MIIERQYJKoZIhvcNAQcCoIIENjCCBDICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUDK6dwcID9g5MGUM0B/eJq2TQ
# JaWgggJPMIICSzCCAbigAwIBAgIQdvBRStRxQLtN+xdsJ3lUIzAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2Vyc2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xNDAxMTUxOTAwMDNaFw0zOTEyMzEyMzU5NTlaMCwxKjAoBgNVBAMTIVJhZmFl
# bCBBbGV4YW5kcmUgRmV1c3RlbCBHdXN0bWFubjCBnzANBgkqhkiG9w0BAQEFAAOB
# jQAwgYkCgYEAmgXb1TwwApRob/zVjgSd6oAUw7YXNWoJRHsqMCAXayQvM9EnlXFs
# CRJwwIEhvXiCH6r1hS/6zrmv9lDt3BEluatXh4H/d5j0tEBooAGgo/XjDmi41Jqa
# vvO5B1HRdEzpOg4frvVhnsePZBeFsQ+hkBEBTc1s+XRzCWAz/KxVKucCAwEAAaN2
# MHQwEwYDVR0lBAwwCgYIKwYBBQUHAwMwXQYDVR0BBFYwVIAQd7PfMUxd/VRgG5cA
# g45ofqEuMCwxKjAoBgNVBAMTIVBvd2Vyc2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUg
# Um9vdIIQuew/eHSS36NIqbMcsQa1IDAJBgUrDgMCHQUAA4GBAAANjbwCEFvAWTXQ
# P6Tixm5blKv/h07STBx2S6bjTPBCmUlhTQD3PdFAvrEniTx8qtdRfLEdNmxKaa26
# 55i7k8NlGPFmCGBFCDEzXM6UyinKOmepTOA2z1Z9byXuUwe284r6Rj4wCTrcDczT
# PgMWCa8pTUjg+0xfrl/MO7yt2PNEMYIBYDCCAVwCAQEwQDAsMSowKAYDVQQDEyFQ
# b3dlcnNoZWxsIExvY2FsIENlcnRpZmljYXRlIFJvb3QCEHbwUUrUcUC7TfsXbCd5
# VCMwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZI
# hvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcC
# ARUwIwYJKoZIhvcNAQkEMRYEFLSV9oA5FNRyzr7APR2N0joNZL+DMA0GCSqGSIb3
# DQEBAQUABIGANU4wZjE50pZ4fT6fr21VB7HyijA9FZijldDL23DQlzDIDQmMCdTZ
# WvdmD2LSwSGExwjJkjvqQkxU/W/Ah5byRepUXdQyoiJj+9J3tvJb9fVo+zusy7b8
# +VtBl6TacnHhIwUzue46smTuEGOLEKr6F696GsG8PpmYskpfR+pkVvw=
# SIG # End signature block
