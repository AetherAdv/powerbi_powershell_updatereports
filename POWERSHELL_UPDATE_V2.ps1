# Connect to Power BI Service Account
Connect-PowerBIServiceAccount

# Set the deployment version to be used for updating parameters
$DeployVer = "2025Q2"

# Define the path to the PBIX file to be used for report updates
$FilePath = "C:\MYFILEPATH\REPORT.pbix"

# Define the conflict action for updating reports (e.g., Create or Overwrite existing reports)
$Conflict = "CreateOrOverwrite"

# Retrieve all Power BI workspaces
$workspaces = Get-PowerBIWorkspace -all

# Loop through each workspace
foreach ($workspace in $workspaces) {

    # Get all reports in the current workspace with names starting with "AETHER" - adjust the filter as needed
    $Reportlist = Get-PowerBIReport -WorkspaceId $workspace.Id | Where-Object -FilterScript {
        $_.Name -LIKE '*AETHER*'
    }

    # Check if any reports were found in the workspace
    if ($Reportlist) {
        Write-Host "Workspace: $($workspace.Name)" # Log the workspace name

        # Loop through each report in the report list
        foreach ($Report in $Reportlist) {
            Write-Host "  Report: $($Report.Name)" # Log the report name

            try {
                # Retrieve the parameters of the dataset associated with the report
                $ParametersJsonString = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/parameters" -Method Get
                $Parameters = (ConvertFrom-Json $ParametersJsonString).value # Convert JSON response to PowerShell object
            } catch {
                Write-Host "Error retrieving parameters: $($_.Exception.Message)"
                continue
            }

            $JsonBase = @{}
            $JsonString = $null # Initialize JSON string variable

            # Initialize an empty array to hold parameter updates
            $UpdateParameterList = New-Object System.Collections.ArrayList

            # Loop through each parameter and prepare the update list
            foreach ($Parameter in $Parameters) {
                $UpdateParameterList.add(@{"name" = $Parameter.name; "newValue" = $Parameter.currentValue})
            }

            # Check if there are any parameters to update
            if ($UpdateParameterList.Count -gt 0) {
                # Get the current value of the Version parameter
                $currentparam = $UpdateParameterList[0].newValue

                Write-Host "Current Parameter Version Value: $currentparam" # Log the current parameter value

                # Check if the current parameter value matches the deployment version
                if ($currentparam -ne $DeployVer) {
                    Write-Host "Version does not match. Updating..." # Log the update action

                    # Display current parameters
                    $UpdateParameterList.newValue

                    # Update the first parameter to the new deployment version
                    $UpdateParameterList[0].newValue = $DeployVer

                    # Prepare the JSON payload for updating parameters
                    $JsonBase.Add("updateDetails", $UpdateParameterList)
                    $JsonString = $JsonBase | ConvertTo-Json

                    # Define the report name
                    $ReportName = $Report.Name

                    # Disable refresh schedule for the dataset
                    $disableRefreshBody = @"
{
"value": {"enabled": false}
}
"@

                    try {
                        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/refreshSchedule" -Method Patch -Body ("$disableRefreshBody")
                        Write-Host "Refresh schedule disabled for dataset: $($Report.DatasetId)"
                    } catch {
                        Write-Host "Failed to disable refresh schedule: $($_.Exception.Message)"
                    }

                    try {
                        # Take over the dataset to ensure permissions are set correctly
                        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/Default.TakeOver" -Method Post
                    } catch {
                        Write-Host "Error taking over dataset: $($_.Exception.Message)"
                        continue
                    }

                    try {
                        # Update the existing report in the workspace
                        New-PowerBIReport -Path $FilePath -Name $ReportName -WorkspaceId $workspace.Id -ConflictAction $Conflict
                    } catch {
                        Write-Host "Error uploading report: $($_.Exception.Message)"
                        continue
                    }

                    try {
                        # Update the parameters of the dataset
                        Start-Sleep 5
                        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/Default.UpdateParameters" -Method Post -Body $JsonString
                    } catch {
                        Write-Host "Error updating parameters: $($_.Exception.Message)"
                        continue
                    }

                    # Reenable refresh schedule for the dataset
                    $enableRefreshBody = @"
{
"value": {"enabled": true}
}
"@

                    try {
                        Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/refreshSchedule" -Method Patch -Body ("$enableRefreshBody")
                        Write-Host "Refresh schedule Enabled for dataset: $($Report.DatasetId)"
                    } catch {
                        Write-Host "Failed to Enable refresh schedule: $($_.Exception.Message)"
                    }

                    Remove-Variable UpdateParameterList, JsonString -ErrorAction SilentlyContinue
                } else {
                    Write-Host "Version already matches. Skipping update." # Log if no update is needed
                }
            } else {
                Write-Host "No parameters found for this dataset." # Log if no parameters are found
            }
        }
    } else {
        Write-Host "No reports found in workspace: $($workspace.Name)" # Log if no reports are found in the workspace
    }
}

# Log the completion of the script
Write-Host "Script completed."