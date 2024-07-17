# Specify your list name and site URL
$ListName = "Incident Management"
$siteUrl = "https://dvagov.sharepoint.com/sites/OperationsTriageGroupOTG2-EngineeringInternal"
$ItemId = 1600

# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get the list and item
$list = Get-PnPList -Identity $ListName
$item = Get-PnPListItem -List $list -Id $ItemId

# Get the item's version history
$versions = Get-PnPProperty -ClientObject $item -Property Versions

# Initialize an array to store version history
$VersionHistory = @()

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

# Export the version history to a CSV file
$VersionHistory | Export-Csv "VersionHistory $($ListName) $($ItemId).csv" -NoTypeInformation