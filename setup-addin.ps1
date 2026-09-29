$ErrorActionPreference = 'Stop'

$pfxPath = Join-Path $PSScriptRoot 'localhost.pfx'
$password = Read-Host 'Enter a password for the local development certificate' -AsSecureString

if (-not (Test-Path $pfxPath)) {
  $certificate = New-SelfSignedCertificate `
    -DnsName 'localhost' `
    -CertStoreLocation 'Cert:\CurrentUser\My' `
    -FriendlyName 'NG Record Add-in development certificate' `
    -Type SSLServerAuthentication `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(2)

  $publicCertificatePath = Join-Path $env:TEMP "ng-record-localhost-$($certificate.Thumbprint).cer"
  try {
    Export-Certificate -Cert $certificate -FilePath $publicCertificatePath | Out-Null
    Import-Certificate -FilePath $publicCertificatePath -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
    Export-PfxCertificate -Cert $certificate -FilePath $pfxPath -Password $password | Out-Null
  } finally {
    Remove-Item -LiteralPath $publicCertificatePath -Force -ErrorAction SilentlyContinue
  }
} else {
  Write-Host 'Using the existing localhost.pfx certificate.'
}

$passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
try {
  $env:NG_ADDIN_CERT_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
} finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
}

Write-Host 'The development certificate is ready for this PowerShell session.'
Write-Host 'Start the local host with: npm start'
