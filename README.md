# **Manage Emails Signature for Outlook in M365**

## **New Outlook**

### Requirements:

1. Ability to install the [ExchangeOnlineManagement Powershell Module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps) and Connect as an Admin
2. An .htm template file of the signature you would like to use
3. An Exported .csv file of the users you want to manage the email signature for
4. **IMPORTANT**: Outlook Roaming Signature must be turned **OFF** for the organization with the Set-OrganizationConfig cmdlet in the ExchangeOnlineManagement Powershell Module beforehand. It may take up to 24 hours for this to take effect. [Set-OrganizationConfig](https://learn.microsoft.com/en-us/powershell/module/exchange/set-organizationconfig?view=exchange-ps#-postponeroamingsignaturesuntillater)

### How does it work?

The script takes a list of users and an .htm template file of a signature and replaces the placeholder values in the template with the actual values in M365/Entra ID for every user. Then, it sets the signature for every user using the [Set-MailBoxMessageConfiguration](https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailboxmessageconfiguration?view=exchange-ps) cmdlet.

### Limitations:

The script only works for the users in the .csv file. Any future users on-boarded will need to have the script run again for them. I recommend adding the .htm template content and the [Set-MailBoxMessageConfiguration](https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailboxmessageconfiguration?view=exchange-ps) cmdlet to your On-Boarding Script.

### How to use:

1. Create an .htm template file of the signature you want to use. You can use a website such as [WordToHTML](https://wordtohtml.net/) to convert your signature to HTML. Replace user properties with placeholder values. Supported placeholder values for the templates are listed below.
2. Either export the users in the M365 Admin portal and delete the rows of users you don't want to manage, or create your own .csv of the users. Just make sure there is a column named **User principal name**
3. Run the SetNewOutllookSignature.ps1 script, using the -csvFilePath and -htmlFilePath parameters

## **Classic Outlook**

### Requirements

1. Windows 10/11 Microsoft Entra ID Joined Devices managed with Microsoft Intune
2. An Exported .csv file of the users you want to manage the email signature for from the [M365 Admin Portal](https://admin.microsoft.com/Adminportal/Home#/users). Click on the ellipses in the top right-hand corner and click **Export users**

### Limitations

- The script only works for users in the .csv file. You will need to repackage the .intunewin file with a new .csv file each time you add new users. A better approach would be to use the ExchangeOnlineManagement Powershell module to grab the user properties instead of a .csv file. ClassicOutlook\Optional\install.ps1 is a script that accomplishes this but I have not been able to get it to work, as the IntuneManagementExtension gets hung up installing the ExchangeOnlineManagement module.

### How to use:

1. Create a signature using the Classic Outlook application and save it. Replace user properties with placeholder values. Supported placeholder values for the templates are listed below.
2. Replace the signature files located in ClassicOutlook\Source\Signatures with the signature files of the signature you just created located in %APPDATA%\Microsoft\Signatures.
3. Replace the ClassicOutlook\Source\ExampleUsers.csv file with the exported .csv from the [M365 Admin Portal](https://admin.microsoft.com/Adminportal/Home#/users)
4. Use the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) to convert the ClassicOutlook\Source folder into an .intunewin file. Use install.ps1 as the setup file.
5. Add a Win32 app in Microsoft Intune using the .intunewin file you just created.
   - Install command (change the name of the .csv file): `Powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "install.ps1" -CsvFilePath .\ExampleUsers.csv`
   - Uninstall command: `PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "uninstall.ps1"`
   - Install behavior: User
   - Return Codes: 0 = Success ; 1 = Failure
   - Detection Rules: Use a custom detection script (ClassicOutlook\detect.ps1)

## Supported Placeholder Values:

_Note: It is important that the actual values are set for the user in M365/Entra ID_

- %City%
- %CountryRegion%
- %Department%
- %DisplayName%
- %FirstName%
- %LastName%
- %MobilePhone%
- %Office%
- %PhoneNumber%
- %State%
- %StreetAddress%
- %Title%
- %UserPrincipalName%
