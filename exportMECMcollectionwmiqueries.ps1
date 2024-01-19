# Function to sanitize the name
function Sanitize-FileName($fileName) {
    $invalidChars = '[\\\/:\*\?"<>\|\[\]]' # Regex to match invalid characters and wildcard characters
    $sanitizedName = $fileName -replace $invalidChars, '_' # Replace invalid characters with underscore
    return $sanitizedName
}

# Load the ConfigurationManager module
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

# Connect to the MECM/SCCM site
$siteCode = "XYZ" # Replace 'XYZ' with your site code
CD "$siteCode`:"

# Path to export the WQL queries
$exportPath = "C:\MECM_CollectionQueries" # Specify the desired export path
if (-not (Test-Path -Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath
}

# Retrieve all device collections
$collections = Get-CMDeviceCollection

foreach ($collection in $collections) {
    # Get the collection rules
    $rules = Get-CMDeviceCollectionQueryMembershipRule -CollectionId $collection.CollectionID

    foreach ($rule in $rules) {
        # Extract the WQL query from the rule
        $wqlQuery = $rule.QueryExpression

        # Sanitize the file name
        $safeCollectionName = Sanitize-FileName $collection.Name
        $safeRuleName = Sanitize-FileName $rule.RuleName

        # Export the WQL query to a text file named after the collection and rule
        $fileName = "$safeCollectionName`_$safeRuleName.txt"
        $filePath = Join-Path -Path $exportPath -ChildPath $fileName
        $wqlQuery | Out-File -FilePath $filePath
    }
}
