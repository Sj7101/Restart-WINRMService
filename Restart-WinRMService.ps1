function Restart-WinRMService {
    param (
        [string]$ComputerName = "localhost"  # Default to localhost if no computer name is specified
    )

    # Function to get the current state of the WinRM service on a remote machine
    function Get-WinRMServiceState {
        try {
            return Get-WmiObject -Class Win32_Service -Filter "Name='WinRM'" -ComputerName $ComputerName
        } catch {
            Write-Host "Error connecting to $ComputerName: $_"
            return $null
        }
    }

    # Function to retrieve the last WinRM service error message from a remote machine
    function Get-ServiceError {
        try {
            $errorLog = Get-WmiObject -Class Win32_NTLogEvent -ComputerName $ComputerName -Filter "Logfile = 'System' AND SourceName = 'Service Control Manager' AND EventType = 1" |
                Where-Object { $_.InsertionStrings -like '*WinRM*' } | Select-Object -First 5
            return $errorLog
        } catch {
            Write-Host "Error fetching error logs from $ComputerName: $_"
            return $null
        }
    }

    # Get the current WinRM service status
    $winrmService = Get-WinRMServiceState

    if ($winrmService) {
        Write-Host "WinRM service found on $ComputerName."

        # Check if the service is running
        if ($winrmService.State -eq 'Running') {
            Write-Host "WinRM service is currently running on $ComputerName. Attempting to restart..."

            # Try to restart the service
            try {
                $result = $winrmService.StopService()  # Stop the service
                Start-Sleep -Seconds 5  # Allow some time for the service to stop

                $result = $winrmService.StartService()  # Start the service
                Start-Sleep -Seconds 5  # Allow some time for the service to start

                # Re-check the service state by querying again
                $winrmService = Get-WinRMServiceState
                if ($winrmService.State -eq 'Running') {
                    Write-Host "WinRM service successfully restarted on $ComputerName."
                } else {
                    Write-Host "WinRM service restart failed on $ComputerName."

                    # Log service error details
                    $errorDetails = Get-ServiceError
                    if ($errorDetails) {
                        $errorDetails | ForEach-Object { 
                            Write-Host "Service Error: $($_.Message)"
                        }
                    }
                }
            } catch {
                Write-Host "Error restarting the WinRM service on $ComputerName: $_"
            }
        } else {
            Write-Host "WinRM service is not running on $ComputerName. Attempting to start..."

            # Try to start the service if it is not running
            try {
                $result = $winrmService.StartService()
                Start-Sleep -Seconds 5  # Allow some time for the service to start

                # Re-check the service state by querying again
                $winrmService = Get-WinRMServiceState
                if ($winrmService.State -eq 'Running') {
                    Write-Host "WinRM service successfully started on $ComputerName."
                } else {
                    Write-Host "Failed to start the WinRM service on $ComputerName."

                    # Log service error details
                    $errorDetails = Get-ServiceError
                    if ($errorDetails) {
                        $errorDetails | ForEach-Object { 
                            Write-Host "Service Error: $($_.Message)"
                        }
                    }
                }
            } catch {
                Write-Host "Error starting the WinRM service on $ComputerName: $_"
            }
        }
    } else {
        Write-Host "WinRM service not found on $ComputerName."
    }
}

foreach($ComputerName in $ComputerNames){
    Restart-WinRMService -ComputerName $ComputerName
}