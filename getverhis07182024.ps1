# Specify your list name and site URL
$ListName = "Incident Management"
$siteUrl = "https://dvagov.sharepoint.com/sites/OperationsTriageGroupOTG2-EngineeringInternal"
$ItemId = 1600

# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get the list
$list = Get-PnPList -Identity $ListName

# Ask the user for input
$choice = Read-Host "Enter '1' to query a specific item or '2' to query items modified in a certain time period"

if ($choice -eq '1') {
    # Query a specific item
    $ItemId = Read-Host "Enter the ID of the item you want to query"
    $items = @(Get-PnPListItem -List $list -Id $ItemId)
} else {
    # Ask the user for the time period
    $timePeriod = Read-Host "Enter the time period in days (24 hours = 1, 7 days = 7, 14 days = 14, 30 days = 30)"
    $date = (Get-Date).AddDays(-$timePeriod)
    $allItems = Get-PnPListItem -List $list
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

                # Check if the field value has changed
                if ($version.FieldValues[$field] -ne $prevVersion.FieldValues[$field]) {

                    # Add change details to the array
                    $VersionHistory += New-Object PSObject -Property @{
                        'INC' = $item.FieldValues["Title"]
                        'ID' = $item.id
                        'VersionId' = $version.VersionId
                        'Created by' = $CreatedBy.Title
                        'Created' = $version.Created
                        'Changed Field' = $field
                        'Old Value' = $prevVersion.FieldValues[$field]
                        'New Value' = $version.FieldValues[$field]
                    }
                }
            }
        }
    }
}

# Export the version history to a CSV file
$VersionHistory | Export-Csv "VersionHistory $($ListName).csv" -NoTypeInformation
