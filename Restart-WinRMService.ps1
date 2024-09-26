function Restart-WinRMService {
    # Define log file location
    $logFile = "G:\Users\Shawn\Desktop\Logs\WinRMServiceRestart.log"
    
    # Log function
    function Log {
        param (
            [string]$message
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp : $message"
        Add-Content -Path $logFile -Value $logMessage
    }

    # Function to get the current state of the WinRM service
    function Get-WinRMServiceState {
        return Get-WmiObject -Class Win32_Service -Filter "Name='WinRM'"
    }

    # Function to retrieve the last WinRM service error message
    function Get-ServiceError {
        $errorLog = Get-EventLog -LogName System -Source "Service Control Manager" -EntryType Error -Newest 5 |
            Where-Object { $_.Message -like "*WinRM*" }
        return $errorLog
    }

    # Get the current WinRM service status
    $winrmService = Get-WinRMServiceState

    if ($winrmService) {
        Log "WinRM service found."

        # Check if the service is running
        if ($winrmService.State -eq 'Running') {
            Log "WinRM service is currently running. Attempting to restart..."

            # Try to restart the service
            try {
                $result = $winrmService.StopService()  # Stop the service
                Start-Sleep -Seconds 5  # Allow some time for the service to stop

                $result = $winrmService.StartService()  # Start the service
                Start-Sleep -Seconds 5  # Allow some time for the service to start

                # Re-check the service state by querying again
                $winrmService = Get-WinRMServiceState
                if ($winrmService.State -eq 'Running') {
                    Log "WinRM service successfully restarted."
                    Write-Host "WinRM service successfully restarted."
                } else {
                    Log "WinRM service restart failed."
                    Write-Host "WinRM service restart failed."

                    # Log service error details
                    $errorDetails = Get-ServiceError
                    if ($errorDetails) {
                        Log "Service Error: $($errorDetails.Message)"
                        Write-Host "Service Error: $($errorDetails.Message)"
                    }
                }
            } catch {
                Log "Error restarting the WinRM service: $_"
                Write-Host "Error restarting the WinRM service: $_"
            }
        } else {
            Log "WinRM service is not running. Attempting to start..."

            # Try to start the service if it is not running
            try {
                $result = $winrmService.StartService()
                Start-Sleep -Seconds 5  # Allow some time for the service to start

                # Re-check the service state by querying again
                $winrmService = Get-WinRMServiceState
                if ($winrmService.State -eq 'Running') {
                    Log "WinRM service successfully started."
                    Write-Host "WinRM service successfully started."
                } else {
                    Log "Failed to start the WinRM service."
                    Write-Host "Failed to start the WinRM service."

                    # Log service error details
                    $errorDetails = Get-ServiceError
                    if ($errorDetails) {
                        Log "Service Error: $($errorDetails.Message)"
                        Write-Host "Service Error: $($errorDetails.Message)"
                    }
                }
            } catch {
                Log "Error starting the WinRM service: $_"
                Write-Host "Error starting the WinRM service: $_"
            }
        }
    } else {
        Log "WinRM service not found on this system."
        Write-Host "WinRM service not found on this system."
    }
}
