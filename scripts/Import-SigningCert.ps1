<#
.SYNOPSIS
	Imports the EXOKit signing certificate so the signed MSIX package can be installed.

.DESCRIPTION
	Imports the publisher's signing certificate into Local Machine Trusted People.
	Verify the certificate fingerprint through your trusted release channel before running.

	Must be run as Administrator.

.PARAMETER CertificatePath
	Path to the EXOKit-Signing-Certificate.cer file. Defaults to a file named
	"EXOKit-Signing-Certificate.cer" in the same folder as this script, which is how it
	ships in the release zip alongside the .msix.

.EXAMPLE
	.\Import-SigningCert.ps1

.EXAMPLE
	.\Import-SigningCert.ps1 -CertificatePath "C:\Downloads\EXOKit-Signing-Certificate.cer"
#>

param(
	[string]$CertificatePath = (Join-Path $PSScriptRoot "EXOKit-Signing-Certificate.cer")
)

if (-not (Test-Path $CertificatePath)) {
	throw "Certificate file not found: $CertificatePath. Pass -CertificatePath if the .cer file is located elsewhere."
}

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
	throw "This script must be run as Administrator (required to write to the Local Machine certificate store)."
}

Write-Host "Importing certificate into Trusted People (Local Machine)..." -ForegroundColor Cyan
Import-Certificate -FilePath $CertificatePath -CertStoreLocation "Cert:\LocalMachine\TrustedPeople" | Out-Null

Write-Host ""
Write-Host "Certificate trusted. You can now install the EXOKit MSIX package." -ForegroundColor Green
