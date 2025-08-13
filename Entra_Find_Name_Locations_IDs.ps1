<#
    .NOTES
    ===========================================================================
    Modified on:
    Created on:
    Created by:   	Alexandros Kapellas
    Organization: 	
    Filename:     	Entra_Find_Name_Locations_IDs.ps1
    ===========================================================================
    ===========================================================================
     Requirements: 
     - Can be run on any machine
    ===========================================================================
    .DESCRIPTION
    This scripts uses MS Graph.
#>

# --- App-only auth (client credentials) ---
$tenantId     = "" # Your TenantID goes here
$clientId     = "" # Your AppID goes here
$clientSecret = "" # Your SecretID goes here

$tok = Invoke-RestMethod -Method POST `
  -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
  -Body @{
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
  }

$jwt = $tok.access_token
$headers = @{ Authorization = "Bearer $jwt" }

# ---- Sanity check: ensure roles are in the token ----
$parts = $jwt.Split('.')
$payloadJson = [Text.Encoding]::UTF8.GetString(
  [Convert]::FromBase64String($parts[1].PadRight($parts[1].Length + (4 - $parts[1].Length % 4) % 4, '=')))
$payload = $payloadJson | ConvertFrom-Json
$roles = @($payload.roles)
"Roles in token: $($roles -join ', ')"
# MUST include: Policy.Read.All  (and for writes: Policy.ReadWrite.ConditionalAccess)

# ---- List Named Locations (v1.0) ----
$resp = Invoke-RestMethod -Method GET -Headers $headers `
  -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?`$top=200"

$resp.value |
  Select-Object id, displayName, @{n='Type';e={$_.PSObject.Properties['@odata.type'].Value}} |
  Format-Table -Auto

# Get the ID for a specific name:
($resp.value | Where-Object displayName -eq "AllowedCountries").id

