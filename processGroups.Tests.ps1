# deleteGroup.Tests.ps1
Describe "deleteGroup Function Tests" {

    BeforeAll {
        # Mock Invoke-RestMethod for all cases
        Mock Invoke-RestMethod {
            if ($Uri -like "https://graph.microsoft.com/v1.0/groups*") {
                return @{
                    value = @(
                        @{
                            displayName = "TestGroup1"
                            id = "group-id-1"
                        },
                        @{
                            displayName = "TestGroup2"
                            id = "group-id-2"
                        }
                    )
                }
            } elseif ($Uri -like "https://graph.microsoft.com/v1.0/groups/group-id-1/members") {
                return @{
                    value = @(
                        @{ id = "user-id-1" },
                        @{ id = "user-id-2" }
                    )
                }
            } elseif ($Method -eq "Delete") {
                return $null
            }
            return $null
        }

        # Mock Update-groupUser to avoid real API calls
        Mock Update-groupUser { Write-Host "User updated: $UserId" }

        # Mock Write-Host and Write-Error for testing
        Mock Write-Host {}
        Mock Write-Error {}
    }

    Context "When deleting an existing group with members" {
        It "Should successfully delete the group and update its members" {
            # Arrange
            $accessToken = "fake-token"
            $groupName = "TestGroup1"

            # Act
            deleteGroup -AccessToken $accessToken -GroupName $groupName

            # Assert
            Should -Invoke Invoke-RestMethod -Times 3 # GET groups, GET members, DELETE group
            Should -Invoke Update-groupUser -Times 2  # For each member in the group
            Should -Invoke Write-Host -ParameterFilter {
                $Object -eq "Group 'TestGroup1' deleted successfully."
            }
        }
    }

    Context "When deleting a non-existent group" {
        It "Should display appropriate message" {
            # Arrange
            $accessToken = "fake-token"
            $groupName = "NonExistentGroup"

            # Act
            deleteGroup -AccessToken $accessToken -GroupName $groupName

            # Assert
            Should -Invoke Write-Host -ParameterFilter {
                $Object -eq "Group 'NonExistentGroup' not found."
            }
        }
    }

    Context "When API call fails" {
        BeforeEach {
            Mock Invoke-RestMethod { throw "API Error" }
        }

        It "Should handle API errors gracefully" {
            # Arrange
            $accessToken = "fake-token"
            $groupName = "TestGroup"

            # Act
            deleteGroup -AccessToken $accessToken -GroupName $groupName

            # Assert
            Should -Invoke Write-Error -ParameterFilter {
                $Message -match "An error occurred while deleting the group"
            }
        }
    }

    Context "Parameter validation" {
        It "Should throw on empty AccessToken" {
            # Act & Assert
            { deleteGroup -AccessToken "" -GroupName "TestGroup" } | 
            Should -Throw -ExpectedMessage "Cannot bind argument to parameter 'AccessToken' because it is an empty string."
        }

        It "Should throw on empty GroupName" {
            # Act & Assert
            { deleteGroup -AccessToken "token" -GroupName "" } | 
            Should -Throw -ExpectedMessage "Cannot bind argument to parameter 'GroupName' because it is an empty string."
        }
    }

    Context "When group has no members" {
        BeforeEach {
            Mock Invoke-RestMethod {
                if ($Uri -like "https://graph.microsoft.com/v1.0/groups/group-id-1/members") {
                    return @{ value = @() } # Return no members
                }
            }
        }

        It "Should delete the group without member updates" {
            # Arrange
            $accessToken = "fake-token"
            $groupName = "TestGroup1"

            # Act
            deleteGroup -AccessToken $accessToken -GroupName $groupName

            # Assert
            Should -Invoke Update-groupUser -Times 0  # No member update calls
        }
    }
}

# Add-UserToAzureADGroup.Tests.ps1]
Describe "Add-UserToAzureADGroup Function Tests" {

    BeforeAll {
        # Mocking the Invoke-RestMethod for all test cases
        Mock Invoke-RestMethod {
            # Return empty response for the POST call
            return $null
        }

        # Mock Log_Message and Error_Log_Message functions
        Mock Log_Message {}
        Mock Error_Log_Message {}
    }

    Context "When adding a user to an existing group" {
        It "Should successfully make the API call and log the success message" {
            # Arrange
            $accessToken = "fake-token"
            $groupId = "group-id-1"
            $userId = "user-id-1"

            # Act
            Add-UserToAzureADGroup -AccessToken $accessToken -GroupId $groupId -UserId $userId

            # Assert
            Should -Invoke Invoke-RestMethod -Exactly 1
            Should -Invoke Log_Message -ParameterFilter {
                $Message -eq "User with ID user-id-1 added to Group with ID group-id-1 successfully."
            }
        }
    }

    Context "When API call fails" {
        BeforeEach {
            Mock Invoke-RestMethod { throw "API Error" }
        }

        It "Should log an error message when the API fails" {
            # Arrange
            $accessToken = "fake-token"
            $groupId = "group-id-1"
            $userId = "user-id-1"

            # Act
            Add-UserToAzureADGroup -AccessToken $accessToken -GroupId $groupId -UserId $userId

            # Assert
            Should -Invoke Error_Log_Message -ParameterFilter {
                $Message -match "An error occurred:"
            }
        }
    }

    Context "Parameter validation" {
        It "Should throw on empty AccessToken" {
            # Act & Assert
            { Add-UserToAzureADGroup -AccessToken "" -GroupName "TestGroup" -UserId "TestUser" } | 
            Should -Throw -ExpectedMessage "Cannot bind argument to parameter 'AccessToken' because it is an empty string."
        }

        It "Should throw on empty GroupId" {
            # Act & Assert
            { Add-UserToAzureADGroup -AccessToken "token" -GroupId "" -UserId "TestUser" } | 
            Should -Throw -ExpectedMessage "Cannot bind argument to parameter 'GroupId' because it is an empty string."
        }

        It "Should throw an error when UserId is missing" {
            { Add-UserToAzureADGroup -AccessToken "fake-token" -GroupId "group-id" -UserId "" } |
            Should -Throw -ExpectedMessage "Cannot bind argument to parameter 'UserId' because it is an empty string."
        }
    }
}

# # Describe block for the Log_Message function
Describe 'Log_Message Function' {

    # Define test-specific variables
    BeforeAll {
        $TestLogFile = "test.log"
        if (-not $TestLogFile) {
            throw "TestLogFile variable is null or empty. Ensure it is initialized properly."
        }

        if (Test-Path $TestLogFile) {
            Remove-Item -Path $TestLogFile -Force
        }

        # Verify that the variable is valid
        if ([string]::IsNullOrWhiteSpace($TestLogFile)) {
            throw "The variable TestLogFile is not properly initialized."
        }

    }

    # Test case: Log message with default log file path
    It 'Should write a log entry to the default file path' {
        # Invoke the function
        Log_Message -Message "Test log entry for default file"

        # Assert that the default log file exists
        $DefaultLogFilePath = "C:\Users\rahulsachin1\Desktop\Powershell\DL_Creation_Logs.txt"
        Test-Path $DefaultLogFilePath | Should -Be $true

        # Assert that the log entry is present
        Get-Content $DefaultLogFilePath | Select-String "Test log entry for default file" | Should -Not -BeNullOrEmpty
    }

    # Test case: Log message with a custom log file path
    It 'Should write a log entry to a specified custom file path' {
        # Invoke the function
        Log_Message -Message "Test log entry for custom file" -LogFilePath $TestLogFile

        # Assert that the custom log file exists
        Test-Path $TestLogFile | Should -Be $true

        # Assert that the log entry is present
        Get-Content $TestLogFile | Select-String "Test log entry for custom file" | Should -Not -BeNullOrEmpty
    }

    # Test case: Log format validation
    It 'Should write the log entry in the correct format' {
        # Invoke the function
        Log_Message -Message "Test log format" -LogFilePath $TestLogFile

        # Read the last log entry
        $LastLogEntry = Get-Content $TestLogFile | Select-Object -Last 1

        # Assert that the log entry contains a timestamp and message
        $LastLogEntry | Should -Match "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} - Test log format$"
    }

    # Cleanup the test log file
    AfterAll {
        if (Test-Path $TestLogFile) {
            Remove-Item -Path $TestLogFile -Force
        }
    }
}
