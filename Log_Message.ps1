# Function to log messages to both the console and the log file
function Log_Message {
    param (
        [string]$Message,
        [string]$LogFilePath = "C:\Users\rahulsachin1\Desktop\Powershell\DL_Creation_Logs.txt" # Default log file path
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
