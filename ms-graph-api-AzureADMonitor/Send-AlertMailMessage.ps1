<#	
	.NOTES
	===========================================================================
	 Created on:   	01/10/2020 11:31
	 Created by:   	Rafael Feustel <esserafael@gmail.com>
	 Filename:     	Send-AlertMailMessage.ps1
	===========================================================================
	.DESCRIPTION
		Sends a fancy alert email message.
#>

Param (
	[Parameter(Mandatory = $true)]
	[String]$To,
	[Parameter(Mandatory = $true)]
	[String]$Title,
	[Parameter(Mandatory = $true)]
	[String]$Content
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

try
{
	$HtmlBody = Get-Content -Raw `
							-Path "$($ScriptPath)\AlertTemplate_inline.html" `
							-Encoding UTF8 `
							-ErrorAction Stop
	
	$HtmlBody = $HtmlBody -replace "{PS_TITLE}", $Title
	$HtmlBody = $HtmlBody -replace "{PS_CONTENT}", $Content
	$HtmlBody = $HtmlBody -replace "{PS_USERNAME}", "$(whoami.exe)"
	$HtmlBody = $HtmlBody -replace "{PS_HOSTNAME}", "$(hostname.exe)"
}
catch
{
	Write-LocalLog "$($_.Exception.Message)"
}

try
{
	Send-MailMessage `
					 -From $MarvinCred.UserName `
					 -To $To `
					 -Subject "[Azure AD Alert] $($Title)" `
					 -Body $HtmlBody `
					 -BodyAsHtml `
					 -Encoding UTF8 `
					 -SmtpServer "smtp.office365.com" `
					 -Port "587" `
					 -UseSsl `
					 -Credential $MarvinCred `
					 -ErrorAction Stop
	
	Write-LocalLog -Text "Mail sent to '$($To)' with title '$($Title)'."
}
catch
{
	Write-LocalLog -Text "Error sending mail: $($_.Exception.Message)"
}
