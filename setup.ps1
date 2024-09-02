$dbstore = "dbstore.csv"
if (test-path $dbstore) {
    write-error "Setup should only be run once. If you want to re-do the setup, delete the 'dbstore.csv' file in this folder"
    Pause
    exit
}
if ($PSVersionTable.psversion.major -lt 6) {
    write-warning "This module is best run in powershell core."
    $powershell = "powershell core"
}
if ($powershell) {
    write-warning "$powershell is not installed, this script module runs best in $powershell"
}
$RSAT = Get-WindowsCapability -Name RSAT* -Online | Select-Object -Property DisplayName, State | where {$_.displayname -eq "RSAT: Active Directory Domain Services and Lightweight Directory Services Tools"}
if ($RSAT.state -eq "NotPresent") {
    Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
}
if($null -eq (get-module -ListAvailable exchangeonlinemanagement)) {
    install-module exchangeonlinemanagement
}
if($null -eq (get-module -ListAvailable microsoft.graph.users)) {
    install-module microsoft.graph.users
}
if($null -eq (get-module -ListAvailable microsoft.graph.applications)) {
    install-module microsoft.graph.applications
}
if($null -eq (get-module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement)) {
    install-module Microsoft.Graph.Identity.DirectoryManagement
}
$certname = "GraphAPI"
$certpath = "$psscriptroot\$certname.cer"
$cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
Export-Certificate -Cert $cert -FilePath $certpath | out-null
$requiredGrants = @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000"
        ResourceAccess = @(
            @{
                Id="38d9df27-64da-44fd-b7c5-a6fbac20248f"
                Type="Role"
            },
            @{
                Id="741f803b-c850-494e-b5df-cde7c675a1ca"
                Type="Role"
            },
            @{
                Id="62a82d76-70ea-41e2-9197-370581804d09"
                Type="Role"
            },
            @{
                Id="19dbc75e-c2e2-444c-a770-ec69d8559fc7"
                Type="Role"
            }
        )
    }
    @{
        ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"
        ResourceAccess = @(
            @{
                Id="dc50a0fb-09a3-484d-be87-e023b12c6440"
                Type="Role"
            }
        )
    }
)
Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read Domain.Read.All Directory.ReadWrite.All RoleManagement.ReadWrite.Directory" -DeviceCode -NoWelcome
$context = Get-MgContext
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
Write-Host -ForegroundColor Cyan "Certificate loaded"
$appRegistration = New-MgApplication -DisplayName "Leavers_process_OnPrem" -SignInAudience "AzureADMyOrg" -Web @{ RedirectUris="http://localhost"; } -RequiredResourceAccess $requiredGrants -AdditionalProperties @{} -KeyCredentials @(@{ Type="AsymmetricX509Cert"; Usage="Verify"; Key=$cert.RawData })
Write-Host -ForegroundColor Cyan "App registration created with app ID" $appRegistration.AppId
$servicePrincipal = New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
$params = @{
	"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($servicePrincipal.id)"
}
New-MgDirectoryRole -RoleTemplateId 29232cdf-9323-42fd-ade2-1d097af3e4de -erroraction SilentlyContinue
$ExchangeRole = Get-MgDirectoryRole -Filter "displayname eq 'Exchange Administrator'"
New-MgDirectoryRoleMemberByRef -DirectoryRoleId $exchangerole.Id -BodyParameter $params
Write-Host -ForegroundColor Cyan "Service principal created"
Write-Host
Write-Host -ForegroundColor Green "Success"
Write-Host
$adminConsentUrl = "https://login.microsoftonline.com/" + $context.TenantId + "/adminconsent?client_id=" + $appRegistration.AppId
Write-Host -ForeGroundColor Yellow "Please go to the following URL in your browser to provide admin consent"
Write-Host $adminConsentUrl
Write-Host
remove-item $certpath -Force
$script:dbdata = @{}
$dbdata["OrgName"] = ((Get-MgOrganization).VerifiedDomains | where {$_.isinitial -eq $true}).name
$dbdata["AppID"] = $appRegistration.AppId
$dbdata['CertThumbprint'] = $cert.Thumbprint
$dbdata['TennantID'] = (Get-MgOrganization).id
$dbdata | export-csv $dbstore
Disconnect-MgGraph | out-null
Write-Host "Disconnected from Microsoft Graph"
