# Function to update user department
function Update-groupUser {
    param (
        [string]$UserId
    )
    $newDepartment = $null

    # Updated property
    $UpdatedProperties = @{
        department = $newDepartment
    }
    # Convert updated properties to JSON
    $Body = $UpdatedProperties | ConvertTo-Json -Depth 10 -Compress

    # Headers
    $Headers = @{
        Authorization = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    # Construct the request URL
    $RequestUrl = "https://graph.microsoft.com/v1.0/users/$UserId"

    # Update the user's department
    try {
        $Response = Invoke-RestMethod -Uri $RequestUrl -Method Patch -Headers $Headers -Body $Body

        # Output success message
        Log_Message "User department updated successfully."
        $Response | ConvertTo-Json -Depth 10 | Write-Output
    } catch {
        Error_Log_Message "Failed to update user's department: $_"
    }
}

# Function to delete group from M365 Groups
function deleteGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,  # OAuth token from MSAL authentication
    
        [Parameter(Mandatory = $true)]
        [string]$GroupName       # Azure AD Group Name to delete
    )
    
    # Validate parameters
    if (-not $AccessToken) {
       throw [System.ArgumentException]::new("AccessToken is required.")
    }
    
    if (-not $GroupName) {
        throw [System.ArgumentException]::new("GroupName is required.")
    }
    
    # Set API endpoint for Microsoft Graph groups
    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/groups"
    
    try {
        # Get the list of groups
        $response = Invoke-RestMethod -Uri $graphApiEndpoint -Headers @{ Authorization = "Bearer $AccessToken" } -Method Get
    
        # Find the group by name
        $group = $response.value | Where-Object { $_.displayName -eq $GroupName }
    
        if ($null -eq $group) {
            Write-Host "Group '$GroupName' not found." -ForegroundColor Red
            return
        }
    
        # Delete the group
        $groupId = $group.id
        # Get the group's members
        $membersEndpoint = "https://graph.microsoft.com/v1.0/groups/$groupId/members"
        $membersResponse = Invoke-RestMethod -Uri $membersEndpoint -Headers @{ Authorization = "Bearer $AccessToken" } -Method Get
        if ($membersResponse.value) {
            $memberIds = $membersResponse.value | ForEach-Object { $_.id }
            $memberIds | ForEach-Object { Update-groupUser -UserId $_ }
        }
        $GraphUri = "https://graph.microsoft.com/v1.0/groups/$groupId"
        Invoke-RestMethod -Uri $GraphUri -Headers @{ Authorization = "Bearer $AccessToken" } -Method Delete
    
        Write-Host "Group '$GroupName' deleted successfully." -ForegroundColor Green
    } catch {
        Write-Error "An error occurred while deleting the group: $_"
    }
}


# Function to Get Access Token (Interactive Login)
function Get-AccessToken {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    # Define the scope for Microsoft Graph
    $Scopes = @("https://graph.microsoft.com/.default")

    # Get the token using MSAL.PS
    $TokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -Scopes $Scopes -Interactive
    return $TokenResponse.AccessToken
}

if($MyInvocation.MyCommand.Path -eq $PSCommandPath){
    # Replace these with your Azure AD app details
    $ClientId = "ed2ecbd5-8bf4-4d34-8085-c66e7f9f36fc"  # Replace with your App Registration Client ID
    $TenantId = "4ab02f96-fd79-417b-9642-9b5fd0a15eeb"  # Replace with your Tenant ID

    # Get Access Token
    $AccessToken = Get-AccessToken -ClientId $ClientId -TenantId $TenantId

    # Replace this with the Group ID you want to delete
    $GroupName = Read-Host "Enter the Group name you want to delete."

    # Call the function to delete the group
    deleteGroup -AccessToken $AccessToken -GroupName $GroupName
}