function Restart-WinRMService {
    param (
        [string]$ComputerName = "localhost"  # Default to localhost if no computer name is specified
    )

    # Function to get the current state of a service on a remote machine
    function Get-ServiceState {
        param (
            [string]$serviceName
        )
        try {
            return Get-WmiObject -Class Win32_Service -Filter "Name='$serviceName'" -ComputerName $ComputerName
        } catch {
            Write-Host "Error connecting to $ComputerName to check $serviceName: $_"
            return $null
        }
    }

    # Function to start a service on a remote machine
    function Start-ServiceRemote {
        param (
            [string]$serviceName
        )
        try {
            $service = Get-ServiceState -serviceName $serviceName
            if ($service -and $service.State -ne 'Running') {
                Write-Host "Attempting to start $serviceName on $ComputerName..."
                $service.StartService()
                Start-Sleep -Seconds 5  # Allow time for the service to start
                $service = Get-ServiceState -serviceName $serviceName  # Re-check the service state
                if ($service.State -eq 'Running') {
                    Write-Host "$serviceName successfully started on $ComputerName."
                } else {
                    Write-Host "Failed to start $serviceName on $ComputerName."
                }
            } elseif ($service -and $service.State -eq 'Running') {
                Write-Host "$serviceName is already running on $ComputerName."
            } else {
                Write-Host "$serviceName not found on $ComputerName."
            }
        } catch {
            Write-Host "Error starting $serviceName on $ComputerName: $_"
        }
    }

    # Function to restart the WinRM service
    function Restart-WinRM {
        # Get the current WinRM service status
        $winrmService = Get-ServiceState -serviceName "WinRM"

        if ($winrmService) {
            Write-Host "WinRM service found on $ComputerName."

            # Check if the service is running
            if ($winrmService.State -eq 'Running') {
                Write-Host "WinRM service is currently running on $ComputerName. Attempting to restart..."

                # Try to restart the service
                try {
                    $winrmService.StopService()  # Stop the service
                    Start-Sleep -Seconds 5  # Allow time for the service to stop

                    $winrmService.StartService()  # Start the service
                    Start-Sleep -Seconds 5  # Allow time for the service to start

                    # Re-check the service state by querying again
                    $winrmService = Get-ServiceState -serviceName "WinRM"
                    if ($winrmService.State -eq 'Running') {
                        Write-Host "WinRM service successfully restarted on $ComputerName."
                    } else {
                        Write-Host "WinRM service restart failed on $ComputerName."
                    }
                } catch {
                    Write-Host "Error restarting the WinRM service on $ComputerName: $_"
                }
            } else {
                Write-Host "WinRM service is not running on $ComputerName. Attempting to start..."

                # Try to start the service if it is not running
                try {
                    $winrmService.StartService()
                    Start-Sleep -Seconds 5  # Allow time for the service to start

                    # Re-check the service state by querying again
                    $winrmService = Get-ServiceState -serviceName "WinRM"
                    if ($winrmService.State -eq 'Running') {
                        Write-Host "WinRM service successfully started on $ComputerName."
                    } else {
                        Write-Host "Failed to start the WinRM service on $ComputerName."
                    }
                } catch {
                    Write-Host "Error starting the WinRM service on $ComputerName: $_"
                }
            }
        } else {
            Write-Host "WinRM service not found on $ComputerName."
        }
    }

    # Step 1: Check and restart RpcSs (Remote Procedure Call) service
    Start-ServiceRemote -serviceName "RpcSs"

    # Step 2: Check and restart HTTP service
    Start-ServiceRemote -serviceName "HTTP"

    # Step 3: Check and restart WinRM service
    Restart-WinRM
}



foreach($ComputerName in $ComputerNames){
    Restart-WinRMService -ComputerName $ComputerName
}