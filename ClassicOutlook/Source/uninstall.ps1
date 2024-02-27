$userPrincipalName = whoami -upn

$signatureFiles = Get-ChildItem -Path "$($env:APPDATA)\Microsoft\Signatures"
foreach ($signatureFile in $signatureFiles) {
    if ($signatureFile.Name -eq "Default Signature ($($userPrincipalName)).txt" -or
        $signatureFile.Name -eq "Default Signature ($($userPrincipalName)).rtf" -or
        $signatureFile.Name -eq "Default Signature ($($userPrincipalName)).htm" -or
        $signatureFile.Name -eq "Default Signature ($($userPrincipalName))_files") {
        Remove-Item -Path $signatureFile.FullName -Recurse -Force
    }
}
