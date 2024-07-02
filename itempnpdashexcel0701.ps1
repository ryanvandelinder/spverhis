# Connect to SharePoint Online
$siteUrl = "https://dvagov.sharepoint.com/sites/OperationsTriageGroupOTG2-EngineeringInternal" # replace with your site URL
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# Get the list
$ListName = "Incident Management" # replace with your list name
$list = Get-PnPList -Identity $ListName

# Get all items
$items = Get-PnPListItem -List $list

# Get the date 30 days ago
$date = (Get-Date).AddDays(-30)

# Filter the items modified in the last 30 days
$recentItems = $items | Where-Object { $_.FieldValues.Modified -gt $date }

# Initialize an array to store change history
$ChangeHistory = @()

# Iterate through each item
foreach ($item in $recentItems) {
    # Get the item's version history
    $versions = Get-PnPProperty -ClientObject $item -Property Versions

    # Iterate through each version
    for ($i=0; $i -lt $versions.Count; $i++) {
        $version = $versions[$i]

        # Get the creator of this version
        $CreatedBy = Get-PnPProperty -ClientObject $version -Property createdby

        # Compare with previous version if it exists
        if ($i -gt 0) {
            $prevVersion = $versions[$i-1]

            # Iterate through each field
            foreach ($field in $version.FieldValues.Keys) {
            #if($field -eq "Primary_x0020_Culpable_x0020_Sys"){
                # Check if the field value has changed
                if ($version.FieldValues[$field].lookupvalue -ne $prevVersion.FieldValues[$field].lookupvalue) {
               # "$version.FieldValues[$field] $prevVersion.FieldValues[$field]"
                $version.FieldValues[$field].lookupvalue
                    # Add change details to the array
                    $ChangeHistory += New-Object PSObject -Property @{
                        'ItemId' = $item.Id
                        'Title' = $item.FieldValues.Title
                        'VersionId' = $version.VersionId
                        'VersionNum' = $version.VersionLabel
                        'Created by' = $CreatedBy.Title
                        'Created' = $version.Created
                        'Changed Field' = $field
                        'New Value' = $version.FieldValues[$field].lookupvalue
                    }
                }
              #  }
            }
        }
    }
}

# Export the change history to a CSV file
$ChangeHistory | Export-Csv -Path "C:\temp\file.csv" -NoTypeInformation

