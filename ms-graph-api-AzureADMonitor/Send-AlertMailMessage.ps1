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
	[String]$Content,
	[Parameter(Mandatory = $true)]
	[String]$AlertSeverity
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
	switch ($AlertSeverity) 
	{
		"medium" { $AlertColor = "#ff9f00" }
		"high" { $AlertColor = "#eb575a" }
		Default { $AlertColor = "#ff9f00" }
	}

	$HtmlBody = Get-Content -Raw `
							-Path "$($ScriptPath)\AlertTemplate_inline.html" `
							-Encoding UTF8 `
							-ErrorAction Stop
	
	$HtmlBody = $HtmlBody -replace "{PS_TITLE}", $Title
	$HtmlBody = $HtmlBody -replace "{PS_ALERTCOLOR}", $AlertColor
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

# SIG # Begin signature block
# MIIERQYJKoZIhvcNAQcCoIIENjCCBDICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUcdJfSKgAOHwyPIChjQ4Zw67Q
# MwKgggJPMIICSzCCAbigAwIBAgIQdvBRStRxQLtN+xdsJ3lUIzAJBgUrDgMCHQUA
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
# ARUwIwYJKoZIhvcNAQkEMRYEFHmakXlmBkcLdrgAmNql1H3Kc1m+MA0GCSqGSIb3
# DQEBAQUABIGAJNgpAT7DEqK1mERQqoiIodPpC3wgXQO2ImFAOE0YGDMZuYodViKx
# GWuzH7pRDb96ruPOXCTrg5jPdy2aMZA7ALvT392GTeIPsAW/hQqpjJKdMlpgkcE7
# ujlK9zlC/fS5EhMfZ+XOgNBVBYTY6QaS7ZvBz/hOxJ0067VdyfIG8hs=
# SIG # End signature block
