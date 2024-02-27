param (
    [Parameter(Mandatory=$true)]
    [string]$CsvFilePath
)

# Import the user information from the CSV file
$users = Import-Csv -Path $CsvFilePath

# Get the UserPrincipalName of the current user
$userPrincipalName = whoami -upn
Write-Host "Current user's UserPrincipalName: $userPrincipalName" -f Blue

# Filter the users based on UserPrincipalName
$user = $users | Where-Object { $_."User principal name" -eq $userPrincipalName }

# If no matching user is found, log an error and exit the script with code 1
if ($null -eq $user) {
    Write-Host "No matching user found in the CSV file for the current user's UserPrincipalName: $userPrincipalName" -f Red
    exit 1 
}

# Check if the signatures folder exists, if not, create it
if (-not (Test-Path "$($env:APPDATA)\Microsoft\Signatures")) {
    Write-Host "Creating signatures folder" -f Blue
    $null = New-Item -Path "$($env:APPDATA)\Microsoft\Signatures" -ItemType Directory
}

# Get all signature files
$signatureFiles = Get-ChildItem -Path "$PSScriptRoot\Signatures"

# Loop through each signature file
foreach ($signatureFile in $signatureFiles) {
    # Check if the file is a .htm, .rtf, or .txt file
    if ($signatureFile.Name -like "*.htm" -or $signatureFile.Name -like "*.rtf" -or $signatureFile.Name -like "*.txt") {
        Write-Host "Processing file: $($signatureFile.Name)" -f Blue

        # Get the content of the file
        $signatureFileContent = Get-Content -Path $signatureFile.FullName

        # Replace the placeholders in the file content with the actual values
        $signatureFileContent = $signatureFileContent -replace "%City%", $user."City"
        $signatureFileContent = $signatureFileContent -replace "%CountryRegion%", $user."Country/Region"
        $signatureFileContent = $signatureFileContent -replace "%Department%", $user."Department"
        $signatureFileContent = $signatureFileContent -replace "%DisplayName%", $user."Display name"
        $signatureFileContent = $signatureFileContent -replace "%FirstName%", $user."First name"
        $signatureFileContent = $signatureFileContent -replace "%LastName%", $user."Last name"
        $signatureFileContent = $signatureFileContent -replace "%MobilePhone%", $user."Mobile Phone"
        $signatureFileContent = $signatureFileContent -replace "%Office%", $user."Office"
        $signatureFileContent = $signatureFileContent -replace "%PhoneNumber%", $user."Phone number"
        $signatureFileContent = $signatureFileContent -replace "%State%", $user."State"
        $signatureFileContent = $signatureFileContent -replace "%StreetAddress%", $user."Street address"
        $signatureFileContent = $signatureFileContent -replace "%Title%", $user."Title"
        $signatureFileContent = $signatureFileContent -replace "%UserPrincipalName%", $user."User principal name"

        # Write the updated content to a new file in the signatures folder
        Set-Content -Path "$($env:APPDATA)\Microsoft\Signatures\Default Signature ($($user."User principal name"))$($signatureFile.Extension)" -Value $signatureFileContent -Force
    } elseif ($signatureFile.getType().Name -eq 'DirectoryInfo') {
        # If the file is a directory, copy it to the signatures folder
        Write-Host "Copying directory: $($signatureFile.Name) to %APPDATA\Microsoft\Signatures\" -f Blue
        Copy-Item -Path $signatureFile.FullName -Destination "$($env:APPDATA)\Microsoft\Signatures\Default Signature ($($user."User principal name"))_files" -Recurse -Force
    }
}

# If the script reaches this point, it has completed successfully
Write-Host "Script completed successfully" -f Green
exit 0
