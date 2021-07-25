<#
.SYNOPSIS
	New Exchange Server Auth Certificate
.DESCRIPTION
	Creates a new Exchange Server Auth (OAuth) Certificate
.EXAMPLE
	New-ExchangeOAuthCertificate
.NOTES
	Author:  Frank Zoechling
	Website: https://www.frankysweb.de
	Twitter: @FrankysWeb
#>

Function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    if ($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    } else {
        return $false
    }
}

#Check for elevated mode
if (-not (Confirm-Administrator)) {
        Write-Warning "This script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator and try again."
        $Error.Clear()
        exit
}

#Create  new OAuth certificate
try {
	write-output "Creating new Exchange OAuth Certificate"
	$SMTPDomainName = (Get-AcceptedDomain | where {$_.default -eq $True}).Domainname.address
	write-output "SMTP Domain Name: $SMTPDomainName"
	$NewOAuthCert = New-ExchangeCertificate -KeySize 2048 -PrivateKeyExportable $true -SubjectName "cn=Microsoft Exchange Server Auth Certificate" -FriendlyName "Microsoft Exchange Server Auth Certificate" -DomainName $SMTPDomainName
	$NewOAuthCertThumbprint = $NewOAuthCert.Thumbprint
	write-output "New Thumbprint: $NewOAuthCertThumbprint"
} catch {
	Write-Error -Message $_.Exception.Message
	exit
}

#Update AuthConfig
try {
	write-output "Updating AuthConfig"
	Set-AuthConfig -NewCertificateThumbprint $NewOAuthCertThumbprint -NewCertificateEffectiveDate (Get-Date)
	Set-AuthConfig -PublishCertificate
	Set-AuthConfig -ClearPreviousCertificate
} catch {
	Write-Error -Message $_.Exception.Message
	exit
}

#Restart IIS AppPools
try {
	write-output "Restaring IISApp Pools"
	Restart-WebAppPool MSExchangeOWAAppPool
	Restart-WebAppPool MSExchangeECPAppPool
} catch {
	Write-Error -Message $_.Exception.Message
	exit
}