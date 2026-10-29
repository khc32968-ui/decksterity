# Finds (or creates) the self-signed code-signing certificate used to sign local Decksterity builds,
# in the current user's personal certificate store, and prints its thumbprint. Used by build.bat.
$subject = 'CN=Decksterity (local build)'
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date) -and $_.HasPrivateKey } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
if (-not $cert) {
    $cert = New-SelfSignedCertificate -Subject $subject -Type CodeSigningCert `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)
}
$cert.Thumbprint
