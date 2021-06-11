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
	[String[]]$CsvPath
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

$ZipFiles = @()

foreach ($Path in $CsvPath)
{
	$DestFile = ($Path -replace ".csv", ".zip")
	$ZipFiles += $DestFile

	$Compress = @{
		Path = $Path
		CompressionLevel = "Optimal"
		DestinationPath = $DestFile
	}
	
	Compress-ReportFiles -Compress $Compress
}

$To = @("paula.rodrigues@uniasselvi.com.br", "pedro.graca@uniasselvi.com.br", "cloves.machado@uniasselvi.com.br")
#$To = @("rafael.gustmann@uniasselvi.com.br")
$DateString = Split-Path $CsvPath -Leaf | Select-String -Pattern "^auditSignIns_(\d+-\d+-\d+)_"

try
{
	Send-MailMessage `
						-From $MarvinCred.UserName `
						-To $To `
						-Cc "rafael.gustmann@uniasselvi.com.br" `
						-Subject "Relatório diário de acessos ao Office 365." `
						-Body "<p style='font-family: 'Segoe UI';'>Olá!<br /><br />Em anexo estão os arquivos com relatórios diários de acessos ao Office 365, do dia $($DateString.Matches.Groups[1].Value).<br /><br />Um arquivo contém os acessos da manhã e tarde (6:00 às 18:00), outro com acessos no período da noite (18:00 à 00:00).</p>" `
						-BodyAsHtml `
						-Encoding UTF8 `
						-Attachments $ZipFiles `
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

$ZipFiles += ($CsvPath -replace ".csv", ".html")

$Compress = @{
	Path = $ZipFiles
	CompressionLevel = "Optimal"
	DestinationPath = ($CsvPath[0] -replace "_Noite\.csv|_Manha\.csv|_Tarde\.csv", "_sources.zip")
}

Compress-ReportFiles -Compress $Compress

try
{
	$CsvPath += ($CsvPath -replace ".csv", ".html")
	Remove-Item -Path $CsvPath -ErrorAction Stop
	Remove-Item -Path $ZipFiles -ErrorAction Stop
}
	catch
{
	Write-LocalLog -Text "Error removing files after compression: $($_.Exception.Message)"
}		
