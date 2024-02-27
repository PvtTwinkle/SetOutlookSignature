Start-Transcript -Path "$($env:TEMP)\IntuneSignatureManagerForOutlook-log.txt" -Force

# Authentication Variables for Connect-ExchangeOnline
$AppId = ""
$CertificateThumbprint = ""
$ORGANIZATION = ""
$CertificateFilePath = ""
$CERTIFICATE_PW = ConvertTo-SecureString -String "" -AsPlainText -Force
$CERTIFICATE_NAME

# Check if PackageManagement module is installed
Write-Host "Checking if PackageManagement module is installed" -f Blue
if (-not (Get-Module -ListAvailable -Name PackageManagement)) {
    Write-Host "PackageManagement module not found, installing..." -f Yellow
    Install-Module -Name PackageManagement -Force -AllowClobber

    # Check if the module is successfully installed
    if (Get-Module -ListAvailable -Name PackageManagement) {
        Write-Host "PackageManagement module installed" -f Green
    } else {
        Write-Host "Failed to install PackageManagement module, exiting..." -f Red
        exit 1
    }
} else {
    Write-Host "PackageManagement module is already installed" -f Green
}

# Check if PowerShellGet module is installed 
Write-Host "Checking if PowerShellGet module is installed" -f Blue
if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
    Write-Host "PowerShellGet module not found, installing..." -f Yellow
    Install-Module -Name PowerShellGet -Force -AllowClobber

    # Check if the module is successfully installed
    if (Get-Module -ListAvailable -Name PowerShellGet) {
        Write-Host "PowerShellGet module installed" -f Green
    } else {
        Write-Host "Failed to install PowerShellGet module, exiting..." -f Red
        exit 1
    }
} else {
    Write-Host "PowerShellGet module is already installed" -f Green
}


# Check if ExchangeOnlineManagement module is installed
Write-Host "Checking if ExchangeOnlineManagement module is installed" -f Blue
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found, installing..." -f Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

    # Check if the module is successfully installed
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "ExchangeOnlineManagement module installed" -f Green
    } else {
        Write-Host "Failed to install ExchangeOnlineManagement module, exiting..." -f Red
        exit 1
    }
} else {
    Write-Host "ExchangeOnlineManagement module is already installed" -f Green
}

Write-Host "Importing ExchangeOnlineManagement module" -f Blue
Import-Module ExchangeOnlineManagement

# Check if the module is successfully imported
if (Get-Module -Name ExchangeOnlineManagement) {
    Write-Host "ExchangeOnlineManagement module imported successfully" -f Green
} else {
    Write-Host "Failed to import ExchangeOnlineManagement module, exiting..." -f Red
    exit 1
}

Write-Host "Importing certificate" -f Blue
import-pfxcertificate -FilePath $CertificateFilePath -Exportable:$false -Password $CERTIFICATE_PW -CertStoreLocation Cert:\CurrentUser\My
$CERTIFICATE = Get-ChildItem cert:\CurrentUser\My\ | where { $_.Subject -eq "cn=$CERTIFICATE_NAME"}
if ($CERTIFICATE) {
    Write-Host "Certicate was successfully imported" -f Green
} else {
    Write-Host "Certificate was not able to be imported, exiting..." -f Red
    exit 1
}

Write-Host "Connecting to Exchange Online" -f Blue
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization

# Test the connection
try {
    Get-AcceptedDomain -ErrorAction Stop > $null
    Write-Host "Connected to Exchange Online successfully." -f Green
}
catch {
    Write-Host "Failed to connect to Exchange Online, exiting" -f Red
    exit 1
}

# Get the user information to update the signature
Write-Host "Getting current user for computer" -f Blue
$userPrincipalName = whoami -upn
if ($userPrincipalName) {
    Write-Host "User $userPrincipalName found" -f Green
} else {
    Write-Host "Failed to find computer's current user, exiting..." -f Red
    exit 1
}
Write-Host "Getting user object information for $userPrincipalName" -f Blue
$userObject = Get-User -Identity $userPrincipalName

# Create signatures folder if it does not exists
if (-not (Test-Path "$($env:APPDATA)\Microsoft\Signatures")) {
    Write-Host "Creating signatures folder" -f Blue
    $null = New-Item -Path "$($env:APPDATA)\Microsoft\Signatures" -ItemType Directory
}

# Get all signature files
$signatureFiles = Get-ChildItem -Path "$PSScriptRoot\Signatures"

foreach ($signatureFile in $signatureFiles) {
    if ($signatureFile.Name -like "*.htm" -or $signatureFile.Name -like "*.rtf" -or $signatureFile.Name -like "*.txt") {
        # Get file content with placeholder values
        $signatureFileContent = Get-Content -Path $signatureFile.FullName

        # Replace placeholder values
        $signatureFileContent = $signatureFileContent -replace "%City%", $userobject.City
        $signatureFileContent = $signatureFileContent -replace "%CountryRegion%", $userObject.CountryOrRegion
        $signatureFileContent = $signatureFileContent -replace "%Department%", $userObject.Department
        $signatureFileContent = $signatureFileContent -replace "%DisplayName%", $userObject.DisplayName
        $signatureFileContent = $signatureFileContent -replace "%FirstName%", $userObject.FirstName
        $signatureFileContent = $signatureFileContent -replace "%LastName%", $userObject.LastName
        $signatureFileContent = $signatureFileContent -replace "%MobilePhone%", $userObject.MobilePhone
        $signatureFileContent = $signatureFileContent -replace "%Office%", $userObject.Office
        $signatureFileContent = $signatureFileContent -replace "%PhoneNumber%", $userObject.Phone
        $signatureFileContent = $signatureFileContent -replace "%State%", $userObject.StateOrProvince
        $signatureFileContent = $signatureFileContent -replace "%StreetAddress%", $userObject.StreetAddress
        $signatureFileContent = $signatureFileContent -replace "%Title%", $userObject.Title
        $signatureFileContent = $signatureFileContent -replace "%UserPrincipalName%", $userObject.UserPrincipalName

        # Set file content with actual values in $env:APPDATA\Microsoft\Signatures
        Set-Content -Path "$($env:APPDATA)\Microsoft\Signatures\Default Signature ($($userPrincipalName))$($signatureFile.Extension)" -Value $signatureFileContent -Force
    } elseif ($signatureFile.getType().Name -eq 'DirectoryInfo') {
        Write-Host "Adding SOFWERX Signature to %APPDATA\Microsoft\Signatures\" -f Blue
        Copy-Item -Path $signatureFile.FullName -Destination "$($env:APPDATA)\Microsoft\Signatures\Default Signature ($($userPrincipalName))_files" -Recurse -Force
    }
}

Write-Host "Removing certificate" -f Blue
Get-ChildItem cert:\CurrentUser\My\ | where { $_.Subject -eq "cn=$CERTIFICATE_NAME"} | Remove-Item

Write-Host "Stopping transcript" -f Blue
Stop-Transcript
exit 0