param (
    [Parameter(Mandatory=$true)]
    [string]$csvFilePath,

    [Parameter(Mandatory=$true)]
    [string]$htmlFilePath
)


# Ensure Required Module is Installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
}

Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Import data from the Excel CSV file for users getting signature
$users = Import-Csv -Path $csvFilePath 

# Iterate through each user
foreach ($user in $users) {
    # Get the user object info 
    $userObject = Get-User -identity $user."User principal name"

    # Get the content of the HTML file
    $htmlFileContent = Get-Content -Path $htmlFilePath

    # Replace the placeholders in the file content with the actual values
    $htmlFileContent = $htmlFileContent -replace "%City%", $userobject.City
    $htmlFileContent = $htmlFileContent -replace "%CountryRegion%", $userObject.CountryOrRegion
    $htmlFileContent = $htmlFileContent -replace "%Department%", $userObject.Department
    $htmlFileContent = $htmlFileContent -replace "%DisplayName%", $userObject.DisplayName
    $htmlFileContent = $htmlFileContent -replace "%FirstName%", $userObject.FirstName
    $htmlFileContent = $htmlFileContent -replace "%LastName%", $userObject.LastName
    $htmlFileContent = $htmlFileContent -replace "%MobilePhone%", $userObject.MobilePhone
    $htmlFileContent = $htmlFileContent -replace "%Office%", $userObject.Office
    $htmlFileContent = $htmlFileContent -replace "%PhoneNumber%", $userObject.Phone
    $htmlFileContent = $htmlFileContent -replace "%State%", $userObject.StateOrProvince
    $htmlFileContent = $htmlFileContent -replace "%StreetAddress%", $userObject.StreetAddress
    $htmlFileContent = $htmlFileContent -replace "%Title%", $userObject.Title
    $htmlFileContent = $htmlFileContent -replace "%UserPrincipalName%", $userObject.UserPrincipalName

    # Convert the file content into a array
    $htmlFileContent = @"
$htmlFileContent
"@

    # Set the email signature for the user
    Set-MailboxMessageConfiguration -identity $userObject.UserPrincipalName -AutoAddSignature:$true -AutoAddSignatureOnReply:$true -SignatureHtml $htmlFileContent
    Write-Host "Signature set for user $($userObject.UserPrincipalName)" -f Green
}
   
# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
