# Azure AD app registration details
$ClientId = "d921e97d-a7cd-4460-9e30-752e0be22ecf"
$TenantId = "8bd725ac-abfd-4696-904c-1baf63cc6ff7"
$ClientSecret = "qSu8Q~zF0k6fIzHAVU23iTuzgUly8gNKPs_JJag_"
# $SecretId = "b98e64c9-97c0-4937-88e1-cffeb0acd4a8"
# Get the access token
$AccessToken = Get-AccessToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret

function Get-AccessToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret
    )

    # Convert the client secret to a SecureString
    $SecureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force

    # Define the scope for Microsoft Graph
    $Scopes = @("https://graph.microsoft.com/.default")

    # Get the token using MSAL.PS with client secret
    $TokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $SecureClientSecret -Scopes $Scopes
    return $TokenResponse.AccessToken
}

$fileId = "48e9e21c-54e1-447f-9e76-03583c304e05"
# $url = "https://graph.microsoft.com/v1.0/me/drive/items/$fileId/workbook/worksheets"

# $response = Invoke-RestMethod -Uri $url -Method GET -Headers @{
#     Authorization = "Bearer $AccessToken"
# }

# $response.value | ForEach-Object {
#     Write-Host "Sheet Name: $($_.name) | Sheet ID: $($_.id)"
# }

$userId = "rahulevrynew_gmail.com#EXT#@rahulevrynewgmail.onmicrosoft.com"  # Replace with the OneDrive user's email or ID
$url = "https://graph.microsoft.com/v1.0/users/$userId/drive/items/$fileId/workbook/worksheets"

$response = Invoke-RestMethod -Uri $url -Method GET -Headers @{
    Authorization = "Bearer $accessToken"
}

$response.value | ForEach-Object {
    Write-Host "Sheet Name: $($_.name) | Sheet ID: $($_.id)"
}
