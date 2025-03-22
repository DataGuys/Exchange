# Step 0: Add the Exchange PowerShell Snap-in if it's not already loaded
if (-not (Get-PSSnapin -Name "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue)) {
    Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn"
}

# Step 1: Choose the SSL certificate to renew
$selectedCert = Get-ExchangeCertificate | Select-Object Subject, CertificateDomains, Thumbprint, NotAfter | Out-GridView -Title "Select a certificate to renew" -PassThru

# Check if a certificate was selected
if ($null -eq $selectedCert) {
    Write-Host "No certificate selected. Exiting..." -ForegroundColor Red
    exit
}

# Show selected certificate's thumbprint
Write-Host "You selected the certificate with thumbprint: $($selectedCert.Thumbprint)" -ForegroundColor Green

# Step 2: Generate CSR for the selected certificate
$txtrequest = Get-ExchangeCertificate -Thumbprint $selectedCert.Thumbprint | New-ExchangeCertificate -GenerateRequest
[System.IO.File]::WriteAllBytes('\\FileServer01\Data\ContosoCertRenewal.req', [System.Text.Encoding]::Unicode.GetBytes($txtrequest))

# Confirm CSR generation
Write-Host "CSR has been generated and saved to '\\FileServer01\Data\ContosoCertRenewal.req'" -ForegroundColor Green

# Get CSR Processed and then save the .cer to a UNC path.

# Step 3: Import the new certificate
$newCertPath = "\\FileServer01\Data\NewCert.cer"
$importedCert = Import-ExchangeCertificate -FileData ([Byte[]]$(Get-Content -Path $newCertPath -Encoding byte -ReadCount 0))

# Enable the new certificate for services
Enable-ExchangeCertificate -Thumbprint $importedCert.Thumbprint -Services "SMTP, IIS, IMAP, POP"

# Confirm Certificate Import and Service Enablement
Write-Host "New certificate has been imported and enabled for services." -ForegroundColor Green
