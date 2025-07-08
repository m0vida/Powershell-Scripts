<#
    .NOTES
    ===========================================================================
    Modified on:
    Created on:
    Created by:   	Alexandros Kapellas
    Organization: 	
    Filename:     	CleanupStaleDevicesEntra.ps1
    ===========================================================================
    ===========================================================================
     Requirements: 
     - Can be run on any machine
    ===========================================================================
    .DESCRIPTION
    This scripts uses MS Graph to find stale devices older than what you specify, and deletes them.
    you can choose to mark out line 45, to test it first, without deleting devices.
#>

# Edit these for your own ORG or comment out if you dont use Certificate for Authentication
$ClientId = "YourAppRegClientID"
$TenantId = "YourTenantID"
$CertificateThumbprint = "CertThumbPrint"

# Connects to Microsoft Graph
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint

# Displays all devices
#$devices = Get-MgDevice -All
#$devices | Select DisplayName, Id, ApproximateLastSignInDateTime


# This command sets the days, that your device will be flagged as "old/staled" Default is 180 days, change accordingly 
$StaleDate = (Get-Date).AddDays(-180).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Get stale devices from Graph
$StaleDevices = Get-MgDevice -Filter "approximateLastSignInDateTime le $StaleDate"

# Displays how many stale devices, that would be deleted
Write-Host "Found $($StaleDevices.Count) stale devices.." -ForegroundColor Yellow

foreach ($device in $StaleDevices) {
    Write-Host "Deleting device: $($device.DisplayName) [$($device.Id)]" -ForegroundColor Cyan
    # Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
}
