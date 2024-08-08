Import-Module ImportExcel
Import-Module PnP.PowerShell



# Specify your list name and site URL
$ListName = "Incident Management"
$siteUrl = "https://dvagov.sharepoint.com/sites/OperationsTriageGroupOTG2-EngineeringInternal"
$ItemId = 1600
Add-Type -AssemblyName System.Windows.Forms

# Function to create a form and return user input
function Get-UserInput ($prompt) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Input Box'
    $form.AutoSize = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(480,40)
    $label.Text = $prompt

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,70)
    $textBox.Size = New-Object System.Drawing.Size(460,20)

    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(380,100)
    $button.Size = New-Object System.Drawing.Size(75,23)
    $button.Text = 'OK'
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.AcceptButton = $button
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($button)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text
    }
}

#external case
#$reportFields = @("test01","test02")

# Array of fields to report on ; if you know of other fields to include , please add to the array below. Values that have been removed Removed:"_UIVersion","_UIVersionString","ContentTypeId",
 
$reportFields = @("Title","Modified","Editor","_IsCurrentVersion","IncidentCloseDate_x002f_Time_x00","IssueStartDate_x002f_TimeisEstim","IssueStartDate_x002f_Time_x0028_","TargetLocationImacted","PrimaryCulpableSystemJustificati",`
"IncidentTimelineNotes","DescriptionofFix","IncidentIssueDescription","NumberofVeteransAffected","NumberofUsersAffected","IncidentState","IncidentImpact","IncdentUrgency","ProjectManager_x002f_SystemOwner","TechnicalPOCEngagedDate_x002f_Ti",`
"Date_x002f_Time_x0028_ET_x0029_R","OTGRecommendations","MitigationActionCategory","MitigationAction","HPI_x002f_PIUpgradeorDowngrade","IncidentSource","IncidentResolvedBy","NameofResolutionProvider","IncidentResolution","IncidentSystemDescription",`
"OTGContributions","TechnicalPOCforIncident","HPI_x002f_CPIReqestDate_x002f_T","IncidentPromotedtoMajorIncident","Date_x002f_Time_x0028_ET_x0029_I","FirstReportedtoESTDate_x002f_Tim","DetectedbyToolNotification","HasMonitoring",`
"IncidentShortDescription","Prmary_x0020_Culpable_x0020_Sys","CausedbyChange","IncidentBusinessImpactDescriptio","Primary_x0020_System_x002f_Appli","CausedbyChangeInfo","IncidentPriority","IncidentSubcategory","IncidentCategor",`
"HowIssuewasReported","Additional_x0020_System_x0028_s_","IncidentPillar","Facilities_x0020_Affected0","EnterpriseFinding_x0028_Architec","Number_x0020_of_x0020_Recommenda","Detected_x0020_by_x020_Config_x","Config_x0020_Mgmt_x0020_Tool_x00",`
"Other_x0020_Facilities_x0020_Aff")

# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get the list
$list = Get-PnPList -Identity $ListName

# Get all items
$allItems = Get-PnPListItem -List $list

# Ask the user if they want to see a list of all items
$listChoice = Get-UserInput "Do you want to see a list of all items? Enter 'yes' or 'no'"

if ($listChoice -eq 'yes') {
    # List all item titles
    Write-Host "Here are the titles of all items in the list:"
    $allItems | ForEach-Object { Write-Host $_.FieldValues["Title"] }
}

# Ask the user for input
$choice = Get-UserInput "Enter '1' to query a specific item or '2' to query items modified in a certain time period"

if ($choice -eq '1') {
    # Query a specific item by title
    $ItemTitle = Get-UserInput "Enter the title of the item you want to query"
    $items = $allItems | Where-Object { $_.FieldValues["Title"] -eq $ItemTitle }
} else {
    # Ask the user for the time period
    $timePeriod = Get-UserInput "Enter the time period in days (24 hours = 1, 7 days = 7, 14 days = 14, 30 days = 30)"
    $date = (Get-Date).AddDays(-$timePeriod)
    $items = $allItems | Where-Object { $_.FieldValues.Modified -gt $date }
}

# Initialize an array to store version history
$VersionHistory = @()

# Iterate through each item
foreach ($item in $items) {
    # Get the item's version history
    $versions = Get-PnPProperty -ClientObject $item -Property Versions

    # Iterate through each version
    for ($i=0; $i -lt $versions.Count; $i++) {
        $version = $versions[$i]
        $CreatedBy = Get-PnPProperty -ClientObject $version -Property createdby

        # Compare with previous version if it exists
        if ($i -gt 0) {
            $prevVersion = $versions[$i-1]

            # Iterate through each field
            foreach ($field in $version.FieldValues.Keys) {
                # Check if the field value has changed and if the field is in the reportFields array
                if ($version.FieldValues[$field] -ne $prevVersion.FieldValues[$field] -and $field -in $reportFields) {
                    # Check if the field value is a user or lookup value
                    if ($version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldUserValue[]] -or $version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldLookupValue[]]) {
                        # Extract the display names of the users or lookup values
                        $newValue = ($version.FieldValues[$field] | ForEach-Object { $_.LookupValue }) -join ', '
                        $oldValue = ($prevVersion.FieldValues[$field] | ForEach-Object { $_.LookupValue }) -join ', '
                    } elseif ($version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldUserValue] -or $version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
                        # Extract the display name of the user or lookup value
                        $newValue = $version.FieldValues[$field].LookupValue
                        $oldValue = $prevVersion.FieldValues[$field].LookupValue
                    } else {
                        # Use the field value directly
                        $newValue = $version.FieldValues[$field]
                        $oldValue = $prevVersion.FieldValues[$field]
                    }

                    # Only add change details to the array if the old value and new value are different
                    if ($oldValue -ne $newValue) {
                        $VersionHistory += New-Object PSObject -Property @{
                            'INC' = $item.FieldValues["Title"]
                            'ID' = $item.id
                            'VersionId' = $version.VersionId
                            'Created by' = $CreatedBy.Title
                            'Created' = $version.Created
                            'Changed Field' = $field
                            'Old Value' = $oldValue
                            'New Value' = $newValue
                        }
                    }
                }
            }
        }
    }
}

# Reorder the VersionHistory
$orderedVersionHistory = $VersionHistory | Select-Object 'INC', 'ID', 'VersionId', 'Created by', 'Created', 'Changed Field', 'Old Value', 'New Value'

# Create a new Excel file path
$excelFilePath = "ChangeFieldAnalysis.xlsx"

# Group data by 'Title' field and create separate worksheets
$groupedData = $orderedVersionHistory | Group-Object -Property 'INC'

foreach ($group in $groupedData) {
    $safeFieldName = $group.Name -replace '[\/:*?"<>|]', '' -replace '\s', '_'
    $worksheetData = $group.Group

    # Initialize a hashtable to count changed field occurrences for the current group
    $changedFieldCounts = @{}
    foreach ($change in $worksheetData) {
        $changedField = $change.'Changed Field'
        $changedFieldCounts[$changedField] = ($changedFieldCounts[$changedField] ?? 0) + 1
    }

    # Add count information to each change in the current group
    $worksheetData = $worksheetData | ForEach-Object {
        New-Object PSObject -Property @{
            'INC' = $_.INC
            'ID' = $_.ID
            'VersionId' = $_.VersionId
            'Created by' = $_.'Created by'
            'Created' = $_.Created
            'Changed Field' = $_.'Changed Field'
            'Old Value' = $_.'Old Value'
            'New Value' = $_.'New Value'
            'Changed Field Count' = $changedFieldCounts[$_.'Changed Field']
        }
    }

    # Export the ordered version history to Excel for the current group
    $worksheetData | Export-Excel -Path $excelFilePath -WorksheetName $safeFieldName -Append -AutoSize

    # Create pivot table for the current worksheet
    $worksheetData | Export-Excel -Path $excelFilePath -WorksheetName $safeFieldName -Append -IncludePivotTable -PivotRows "Changed Field" -PivotData @{ 'Changed Field Count' = 'Sum' } -PivotTableName "PivotTable$safeFieldName"

    # Add pivot chart for the current worksheet
    $worksheetData | Export-Excel -Path $excelFilePath -WorksheetName $safeFieldName -Append -IncludePivotChart -ChartType ColumnClustered -PivotTableName "PivotTable$safeFieldName"
}

Write-Host "Analysis complete. Excel file saved as $excelFilePath"
