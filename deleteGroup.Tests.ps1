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