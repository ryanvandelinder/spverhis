# Specify your list name and site URL
$ListName = "Incident Management"
$siteUrl = "https://dvagov.sharepoint.com/sites/OperationsTriageGroupOTG2-EngineeringInternal"
$ItemId = 1600


Add-Type -AssemblyName System.Windows.Forms

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

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $textBox.Text
    }
}

#external case
#$reportFields = @("test01","test02")

# Array of fields to report on ; if you know of other fields to include , please add to the array below 
$reportFields = @(<#"Modified",#>"ProjectManager_x002f_SystemOwner","TechnicalPOCforIncident","NameofResolutionProvider","Facilities_x0020_Affected0","Primary_x0020_Culpable_x0020_Sys","Primary_x0020_System_x002f_Appli",`
"Number_x0020_of_x0020_Recommenda","HasMonitoring","TargetLocationImpacted","IncidentTimelineNotes","IncidentPromotedtoMajorIncident","IncidentPillar","HPI_x002f_CPIUpgradeorDowngradeFirstReportedtoESTDate_x002f_Tim",`
"EnterpriseFinding_x0028_Architec","Detected_x0020_by_x0020_Config_x","Date_x002f_Time_x0028_ET_x0029_I","Config_x0020_Mgmt_x0020_Tool_x00")

Connect-PnPOnline -Url $siteUrl -UseWebLogin

$list = Get-PnPList -Identity $ListName

$listChoice = Get-UserInput "Do you want to see a list of all items? Enter 'yes' or 'no'"

if ($listChoice -eq 'yes') {
    $allItems = Get-PnPListItem -List $list
    Write-Host "Here are the titles of all items in the list:"
    $allItems | ForEach-Object { Write-Host $_.FieldValues["Title"] }
}

$choice = Get-UserInput "Enter '1' to query a specific item or '2' to query items modified in a certain time period"

if ($choice -eq '1') {
    $ItemTitle = Get-UserInput "Enter the title of the item you want to query"
    $items = Get-PnPListItem -List $list | Where-Object { $_.FieldValues["Title"] -eq $ItemTitle }
} else {
    $timePeriod = Get-UserInput "Enter the time period in days (24 hours = 1, 7 days = 7, 14 days = 14, 30 days = 30)"
    $date = (Get-Date).AddDays(-$timePeriod)
    $items = Get-PnPListItem -List $list | Where-Object { $_.FieldValues.Modified -gt $date }
}

$VersionHistory = @()

foreach ($item in $items) {
    $versions = Get-PnPProperty -ClientObject $item -Property Versions

    for ($i=0; $i -lt $versions.Count; $i++) {
        $version = $versions[$i]
        $CreatedBy = Get-PnPProperty -ClientObject $version -Property createdby

        if ($i -gt 0) {
            $prevVersion = $versions[$i-1]

            foreach ($field in $version.FieldValues.Keys) {
                if ($version.FieldValues[$field] -ne $prevVersion.FieldValues[$field] -and $field -in $reportFields) {
                    if ($version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldUserValue[]] -or $version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldLookupValue[]]) {
                        $oldValue = ($version.FieldValues[$field] | ForEach-Object { $_.LookupValue }) -join ', '
                        $newValue = ($prevVersion.FieldValues[$field] | ForEach-Object { $_.LookupValue }) -join ', '
                    } elseif ($version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldUserValue] -or $version.FieldValues[$field] -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
                        $oldValue = $version.FieldValues[$field].LookupValue
                        $newValue = $prevVersion.FieldValues[$field].LookupValue
                    } else {
                        $oldValue = $version.FieldValues[$field]
                        $newValue = $prevVersion.FieldValues[$field]
                    }

                    # Adjust for time zone difference
                    if ($oldValue -is [DateTime]) {
                        $oldValue = $oldValue.AddHours(-4)
                    }
                    if ($newValue -is [DateTime]) {
                        $newValue = $newValue.AddHours(-4)
                    }

                    if ($oldValue -ne $newValue) {
                        $VersionHistory += New-Object PSObject -Property @{
                            'INC' = $item.FieldValues["Title"]
                            'ID' = $item.id
                            'VersionId' = $version.VersionId
                            'Created by' = $CreatedBy.Title
                            'Created' = $version.Created.AddHours(-4) # Adjust for time zone difference
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

$orderedVersionHistory = $VersionHistory | Select-Object 'INC', 'ID', 'VersionId', 'Created by', 'Created', 'Changed Field', 'Old Value', 'New Value'

$orderedVersionHistory | Export-Csv "VersionHistory $($ListName).csv" -NoTypeInformation


