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
