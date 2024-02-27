$userPrincipalName = whoami -upn
$signatureFiles = Get-ChildItem -Path "$($env:APPDATA)\Microsoft\Signatures"

# Define the files/folders to check
$itemsToCheck = @(
    "Default Signature ($($userPrincipalName)).txt",
    "Default Signature ($($userPrincipalName)).rtf",
    "Default Signature ($($userPrincipalName)).htm",
    "Default Signature ($($userPrincipalName))_files"
)

# Check if each item exists in the signatureFiles array
foreach ($item in $itemsToCheck) {
    if ($signatureFiles.Name -notcontains $item) {
        # If the item does not exist, exit with "1"
        Write-Host "User does not have the Default Signature" -f Red
        exit 1
    }
}

# If all items exist, exit with "0"
Write-Host "User has the Default Signature already" -f Green
exit 0
