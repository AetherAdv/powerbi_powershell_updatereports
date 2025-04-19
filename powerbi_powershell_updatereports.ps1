# Connect to Power BI Service Account
Connect-PowerBIServiceAccount

# Set the deployment version to be used for updating parameters
$DeployVer = "2025Q1"

# Define the path to the PBIX file to be used for report updates
$FilePath = "C:\Temp\myreport.pbix"

# Define the conflict action for updating reports (e.g., Create or Overwrite existing reports)
$Conflict = "CreateOrOverwrite"

# Retrieve all Power BI workspaces
$workspaces = Get-PowerBIWorkspace -All

# Loop through each workspace
foreach ($workspace in $workspaces) {

    # Get all reports in the current workspace with names starting with "AETHER"
    $Reportlist = Get-PowerBIReport -WorkspaceId $workspace.Id | Where-Object { $_.Name -like 'AETHER*' }

    # Check if any reports were found in the workspace
    if ($Reportlist) {
        Write-Host "Workspace: $($workspace.Name)" # Log the workspace name

        # Loop through each report in the report list
        foreach ($Report in $Reportlist) {
            Write-Host "  Report: $($Report.Name)" # Log the report name

            $JsonString = $null # Initialize JSON string variable

            # Retrieve the parameters of the dataset associated with the report
            $ParametersJsonString = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/parameters" -Method Get
            $Parameters = (ConvertFrom-Json $ParametersJsonString).value # Convert JSON response to PowerShell object

            # Initialize an empty array to hold parameter updates
            $UpdateParameterList = @()

            # Loop through each parameter and prepare the update list
            foreach ($Parameter in $Parameters) {
                $UpdateParameterList += @{ "name" = $Parameter.name; "newValue" = $Parameter.currentValue }
            }

            # Check if there are any parameters to update
            if ($UpdateParameterList.Count -gt 0) {
                # Get the current value of the first parameter
                $currentparam = $UpdateParameterList[0].newValue

                Write-Host "    Current Parameter 0 Value: $currentparam" # Log the current parameter value

                # Check if the current parameter value matches the deployment version
                if ($currentparam -ne $DeployVer) {
                    Write-Host "Version does not match. Updating..." # Log the update action

                    # Update the first parameter to the new deployment version
                    $UpdateParameterList[0].newValue = $DeployVer
                }
                else {
                    Write-Host "Version already matches. Skipping update." # Log if no update is needed
                }

                # Prepare the JSON payload for updating parameters
                $JsonBase = @{ "updateDetails" = $UpdateParameterList }
                $JsonString = $JsonBase | ConvertTo-Json

                # Define the report name
                $ReportName = $Report.Name

                # Update the existing report in the workspace
                New-PowerBIReport -Path $FilePath -Name $ReportName -WorkspaceId $workspace.Id -ConflictAction $Conflict

                # Take over the dataset to ensure permissions are set correctly
                Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/Default.TakeOver" -Method Post

                # Update the parameters of the dataset
                Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/Default.UpdateParameters" -Method Post -Body $JsonString

                # Pause for 5 seconds to avoid API rate limits
                Start-Sleep -Seconds 5

                # Trigger a dataset refresh
                Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($Report.DatasetId)/refreshes" -Method Post
                Write-Host "Refresh started." # Log the refresh action
            }
            else {
                Write-Host "No parameters found for this dataset." # Log if no parameters are found
            }
        }
    } else {
        Write-Host "No reports found in workspace: $($workspace.Name)" # Log if no reports are found in the workspace
    }
}

# Log the completion of the script
Write-Host "Script completed."