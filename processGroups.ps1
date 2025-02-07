# Import Excel Module
Import-Module ImportExcel -ErrorAction Stop

# Azure AD app registration details
$ClientId = "**********************************"
$TenantId = "**********************************"
$ClientSecret = "**********************************"
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

$ErrorFilePath =  "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\DL_Creation_ErrorLogs.txt"

function Log_Message {
    param (
        [string]$Message,
        [string]$LogFilePath = "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\DL_Creation_Logs.txt" # Default log file path
    )
    # Get the current timestamp
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Format the log entry
    $LogEntry = "$Timestamp - $Message"
    # Write to console
    Write-Host $LogEntry
    # Append to log file
    Add-Content -Path $LogFilePath -Value $LogEntry
}

function Error_Log_Message {
    param (
        [string]$Message
    )
    # Get the current timestamp
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Format the log entry
    $LogEntry = "$Timestamp - $Message"
    # Append to log file
    Add-Content -Path $ErrorFilePath -Value $LogEntry
}

# Function to get Azure AD Users using Graph API
function Get-AzureADUsers {
    param (
        [string]$AccessToken
    )

    $graphApiUrl = "https://graph.microsoft.com/v1.0/users?`$select=displayName,userPrincipalName,id,department,jobTitle&`$top=999"
    $response = Invoke-RestMethod -Uri $graphApiUrl -Headers @{ Authorization = "Bearer $AccessToken" } -Method Get
    return $response.value
}

function deleteUser {
    param (
        [string]$UserId
    )
    # Headers
    $Headers = @{
        Authorization = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    # Construct the request URL
    $RequestUrl = "https://graph.microsoft.com/v1.0/users/$UserId"

    # Delete the user
    try {
        Invoke-RestMethod -Uri $RequestUrl -Method Delete -Headers $Headers

        # Output success message
        Log_Message "User with ID '$UserId' deleted successfully."
    } catch {
        Error_Log_Message "Failed to delete user: $_"
    }
    
}

function restoreUser {
    param (
        [string]$UserId
    )
    # Restore the deleted user by userId
    $restoreUrl = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user/$UserId/restore"
    try {
        $response = Invoke-RestMethod -Method Post -Uri $restoreUrl -Headers @{
            Authorization = "Bearer $AccessToken"
            "Content-Type" = "application/json"
        }
        Log_Message "User restored successfully! User ID: $($response.id)" -ForegroundColor Green
    } catch {
        Error_Log_Message "Error restoring user: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to update any existing User
function updateUser {
    param (
        [string]$userId,
        [string]$newdisplayName,
        [string]$newjobTitle,
        [string]$newDepartmentName
    )
    if($newDepartmentName -eq ""){
        $newDepartment = $null
    }
    if($newjobTitle -eq ""){
        $newJob = $null
    }
    # Updated property
    $UpdatedProperties = @{
        displayName = $newdisplayName
        jobTitle = $newJob
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

function Update-AzureAD{
    # Define the Excel file path
    $excelFilePath = "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\UserData.xlsx"

    # Get Azure AD users
    $azureADUsers = Get-AzureADUsers -AccessToken $AccessToken

    $excelData = Import-Excel -Path $excelFilePath -WorksheetName "AzureADUsers"

    foreach ($user in $excelData){
        
        $matchingUser = $azureADUsers | Where-Object { $_.id -eq $user.id }
        if($user.is_active -eq 0){
            if($matchingUser){
                deleteUser -UserId $($user.id)
                continue
            }else{
                continue
            }
        }
        if ($matchingUser) {
            $newDepartment = $user.department
            updateUser -UserId $($user.id) -newdisplayName $($user.displayName) -newjobTitle $($user.jobTitle) -newDepartmentName $newDepartment
        }else{
            if(($user.id).Length -eq 36){
                restoreUser -UserId $($user.id)
                $newDepartment = $user.department
                updateUser -UserId $($user.id) -newdisplayName $($user.displayName) -newjobTitle $($user.jobTitle) -newDepartmentName $newDepartment
            }else{
                Log_Message "Add a new User."
            }
        }
    }
}

function Update-Excel{
    # Define the Excel file path
    $excelFilePath = "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\UserData.xlsx"
    # Get Azure AD users
    $azureADUsers = Get-AzureADUsers -AccessToken $AccessToken

    # Read the Excel file
    if (-Not (Test-Path $excelFilePath)) {
        Log_Message "Excel file not found. Creating a new file."
        $excelData = @()
        Export-Excel -Path $excelFilePath -WorksheetName "AzureADUsers" -Data $excelData -AutoSize
    }

    $excelData = Import-Excel -Path $excelFilePath -WorksheetName "AzureADUsers"

    # Check for the 'is_active' column in the Excel data
    if (-Not ($excelData -and $excelData[0].PSObject.Properties.Name -contains 'is_active')) {
        $excelData | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "is_active" -Value $null
        }
    }

    # Process the users
    $updatedData = @()

    foreach ($user in $azureADUsers) {
        $matchingUser = $excelData | Where-Object { $_.id -eq $user.id }
        if ($matchingUser) {
            # Update is_active column for existing users
            $matchingUser.is_active = 1
            $matchingUser.displayName = $user.displayName
            if($null -eq $user.department){
                $matchingUser.department = ""
            }else{
                $matchingUser.department = $user.department
            }
            $updatedData += $matchingUser
        } else {
            # Add new user
            $updatedData += [PSCustomObject]@{
                id                = $user.id
                displayName       = $user.displayName
                userPrincipalName = $user.userPrincipalName
                department        = $user.department
                jobTitle          = $user.jobTitle
                is_active         = 1
            }
        }
    }

    foreach ($user in $excelData){
        $matchingUser = $azureADUsers | Where-Object { $_.id -eq $user.id }
        if ($matchingUser) {
            continue
        }else{
            $user.is_active = 0
            $updatedData += $user
        }
    }

    # Export the sorted data back to the Excel file
    $sortedData = $updatedData | Sort-Object -Property displayName
    $sortedData | Export-Excel -Path $excelFilePath -WorksheetName "AzureADUsers" -AutoSize

    Log_Message "User details have been updated in the Excel file: $excelFilePath"
}

function Get-ExcelLastUpdatedTime {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath  # Path to the Excel file
    )

    try {
        # Check if the file exists
        if (-not (Test-Path -Path $FilePath)) {
            throw [System.IO.FileNotFoundException]::new("File not found: $FilePath")
        }

        # Get the file properties
        $file = Get-Item -Path $FilePath

        # Retrieve the LastWriteTime property
        $lastUpdatedTime = $file.LastWriteTime

        # Format the date and time as MM/dd/yyyy HH:mm:ss
        $formattedTime = $lastUpdatedTime.ToString("MM'/'dd'/'yyyy HH:mm:ss")
        return $formattedTime
    }
    catch {
        # Handle errors
        Write-Error "An error occurred: $_"
    }
}

function auditLogs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,  # Access Token
        
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName  # User Principal Name to filter logs
    )
    
    # Base URL for audit logs
    $baseUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits"

    # Define query parameters
    $queryParams = "?`$filter=result eq 'Success' and initiatedBy/user/userPrincipalName eq '$UserPrincipalName' and category eq 'UserManagement'&`$orderby=activityDateTime desc"

    # Full URL with query parameters
    $uri = "$baseUri$queryParams"

    # Define headers with the access token
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    # Make the API call
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ContentType "application/json"
        
        if ($response.value) {
            foreach ($log in $response.value) {
                $time = $($log.activityDateTime)
                $actualTime = $time.ToLocalTime()
                return $actualTime
            }
        } else {
            Log_Message "No audit logs found for the specified user with 'Success' status."
        }
    } catch {
        Error_Log_Message "Error occurred while fetching audit logs: $_"
    }
}

function Get-GraphGroups {
    param (
        [string]$AccessToken
    )

    # Set API endpoint for Microsoft Graph groups
    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/groups"

    # Initialize an array to store group details
    $groupDetails = @()

    # Fetch groups using the Microsoft Graph API
    try {
        $response = Invoke-RestMethod -Uri $graphApiEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
        
        # Process each group and store the display name and ID
        foreach ($group in $response.value) {
            $groupDetails += [PSCustomObject]@{
                DisplayName = $group.displayName
                GroupId     = $group.id
            }
        }

        # Store group details in a variable for further use
        $GroupsData = $groupDetails
        Log_Message "Group details successfully stored in variable 'GroupsData'"
        return $GroupsData
    } catch {
        Error_Log_Message "An error occurred: $_"
        return $null
    }
}

function NewGroup {
    param (
        [string]$AccessToken,
        [hashtable]$GroupDetails
    )

    # Set API endpoint for creating the group
    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/groups"

    # Convert group details to JSON
    $groupDetailsJson = $GroupDetails | ConvertTo-Json -Depth 10 -Compress

    # Create the group using the Microsoft Graph API
    try {
        $response = Invoke-RestMethod -Uri $graphApiEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Body $groupDetailsJson -ContentType "application/json" -Method Post
        
        # Output the created group's details
        Log_Message "Group created successfully!"
        Log_Message "Group ID: $($response.id)"
        Log_Message "Group Display Name: $($response.displayName)"
        Log_Message "Group Mail Nickname: $($response.mailNickname)"
        return $response
    } catch {
        Error_Log_Message "An error occurred: $_"
        return $null
    }
}

# Function to get Group ID by department name
function Get-GroupIdByName {
    param (
        [string]$departmentName,
        [array]$MatchingData
    )

    foreach ($entry in $MatchingData) {
        if ($entry.Department -eq $DepartmentName) {
            return $entry.GroupId
        }
    }
    return $null # Return null if the department is not found
}

# Function to add a user to a group in Azure AD
function Add-UserToAzureADGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,  # OAuth token from MSAL authentication

        [Parameter(Mandatory = $true)]
        [string]$GroupId,      # Azure AD Group ID

        [Parameter(Mandatory = $true)]
        [string]$UserId        # Azure AD User ID
    )
    # Microsoft Graph API endpoint to add a member to a group
    $GraphUri = "https://graph.microsoft.com/v1.0/groups/{$GroupId}/members/`$ref"

    # Define the request body
    $Body = @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/{$UserId}"
    } | ConvertTo-Json -Depth 10

    # Set the headers with the access token
    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    # Make the API call
    try {
        Invoke-RestMethod -Method Post -Uri $GraphUri -Body $Body -Headers $Headers
        Log_Message "User with ID $UserId added to Group with ID $GroupId successfully."
    } catch {
        Error_Log_Message "An error occurred: $_"
    }
}

# Function to check if user is in the group
function UserInGroup {
    param (
        [string]$AccessToken,
        [string]$UserId,
        [string]$GroupId
    )

    # API Endpoint
    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"

    try {
        $response = Invoke-RestMethod -Uri $graphApiEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
        foreach ($member in $response.value) {
            if ($member.id -eq $UserId) {
                return $true
            }
        }
        return $false
    } catch {
        Error_Log_Message "An error occurred while checking user membership: $_"
        return $false
    }
}

function Remove-UserFromGroup {
    param (
        [string]$AccessToken,  # Access token for authentication
        [string]$UserId,       # ID of the user to remove
        [string]$GroupId       # ID of the group to remove the user from
    )

    # Define the URI for the API request to remove the user from the group
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$UserId/`$ref"

    # Send the DELETE request to remove the user from the group
    try {
        Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"} -Method Delete
        Log_Message "Successfully removed user $UserId from group $GroupId."
    } catch {
    }
}

function Remove-UserFromGroupIfNotMatched {
    param (
        [string]$AccessToken,
        [string]$UserId,
        [string]$GroupId,
        [array]$MatchingData  # Array containing matching group IDs
    )
    foreach ($entry in $MatchingData){
        # Check if the GroupId matches any entry in the MatchingData array
        $currGroupID = $entry.GroupId
        if ($currGroupID -eq $GroupId) {
            Log_Message "GroupId $GroupId matches an entry in the MatchingData array. No action needed."
        } else {
            Remove-UserFromGroup -AccessToken $AccessToken -UserId $UserId -GroupId $currGroupID
        }
    }
}

function Process_DistributionLists {
    Log_Message "Processing DL"
    # Define the file path for the Excel data
    $ExcelFilePath = "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\UserData.xlsx"

    # Import the data from the Excel file
    $EmployeeData = Import-Excel -Path $ExcelFilePath

    # Get all the department and member details from the excel
    $Departments = $EmployeeData | Select-Object -ExpandProperty department | Sort-Object -Unique

    # Initialize $MatchingData as an empty array
    $MatchingData = @()

    foreach ($Department in $Departments) {
        $GroupsData = Get-GraphGroups -AccessToken $AccessToken
        $GroupFound = $false

        foreach ($Group in $GroupsData) {
            if ($Group.DisplayName -like "*$Department*") {
                $MatchingData += [PSCustomObject]@{
                    Department = $Department
                    GroupName  = $Group.DisplayName
                    GroupId    = $Group.GroupId
                }
                $GroupFound = $true
                break
            }
        }

        if (-not $GroupFound) {
            $NewGroupDetails = @{
                displayName     = $Department
                mailNickname    = $Department.ToLower() -replace "\s", ""
                mailEnabled     = $true
                securityEnabled = $false
                groupTypes      = @("Unified")
            }
            $NewGroupResponse = NewGroup -AccessToken $AccessToken -GroupDetails $NewGroupDetails
            if ($NewGroupResponse) {
                $MatchingData += [PSCustomObject]@{
                    Department = $Department
                    GroupName  = $NewGroupResponse.displayName
                    GroupId    = $NewGroupResponse.id
                }
            }
        }
    }

    Log_Message $MatchingData

    # Iterate through the Excel data
    foreach ($row in $EmployeeData) {
        if($row.is_active -eq 1){
            $UserId = $row.id
            $departmentName = $row.department

            # Check if departmentName is null or empty before calling the function
            if (![string]::IsNullOrEmpty($departmentName)) {
                # Get group ID
                $GroupId = Get-GroupIdByName -AccessToken $AccessToken -departmentName $departmentName -MatchingData $MatchingData
                #Log_Message $GroupId
                if ($GroupId) {
                    Log_Message "User $UserId needs to be added to group $GroupId"
                    # Check if the user is in the group
                    $isInGroup = UserInGroup -AccessToken $AccessToken -UserId $UserId -GroupId $GroupId
                    if (-not $isInGroup) {
                        # Add the user to the group
                        Add-UserToAzureADGroup -AccessToken $AccessToken -GroupId $GroupId -UserId $UserId
                        # Call the function
                        Remove-UserFromGroupIfNotMatched -AccessToken $AccessToken -UserId $UserId -GroupId $GroupId -MatchingData $MatchingData
                    }
                }
            } else {
                Log_Message "Department is null or empty for user $UserId"
                foreach ($entry in $MatchingData){
                    # Check if the GroupId matches any entry in the MatchingData array
                    $currGroupID = $entry.GroupId
                    Remove-UserFromGroup -AccessToken $AccessToken -UserId $UserId -GroupId $currGroupID
                }
            }
        }
    }
}

# Infinite loop to run the script every 10 minutes
while ($true) {
    try {
        # Call the auditLogs function with the desired user and access token
        $azureADtime = (auditLogs -AccessToken $AccessToken -UserPrincipalName "sachinrahul@sachinrahul.onmicrosoft.com")

        $excelFilePath = "C:\Users\rahulsachin1\Desktop\Hospice Final Presentation\UserData.xlsx"
        $excelTime = [datetime](Get-ExcelLastUpdatedTime -FilePath $excelFilePath)
        
        if($excelTime -gt $azureADtime){
            Log_Message "Excel is updated now update Azure AD"
            Update-AzureAD
            Process_DistributionLists
        }else{
            Log_Message "Azure AD is updated last."
            Update-Excel
            Process_DistributionLists
        }
    } catch {
        Error_Log_Message "An error occurred: $_"
    }
    Log_Message "Script execution completed. Waiting for the next run."
    Start-Sleep -Seconds 60 # 1 minutes
}
