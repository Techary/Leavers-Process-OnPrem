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
if($null -eq (get-module -ListAvailable microsoft.graph)) {
    install-module microsoft.graph
}
$certname = "GraphAPI"
$certpath = "$psscriptroot\$certname.cer"
$cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
Export-Certificate -Cert $cert -FilePath $certpath
$graphResourceId = "00000003-0000-0000-c000-000000000000"
$UserAuthenticationMethodReadAll = @{
    Id="38d9df27-64da-44fd-b7c5-a6fbac20248f"
    Type="Role"
}
$UserReadWriteAll = @{
    Id="741f803b-c850-494e-b5df-cde7c675a1ca"
    Type="Role"
}
$GroupReadWriteAll = @{
    Id="62a82d76-70ea-41e2-9197-370581804d09"
    Type="Role"
}
$DirectoryReadWriteAll = @{
    Id="19dbc75e-c2e2-444c-a770-ec69d8559fc7"
    Type="Role"
}
$ExchangeManageAsApp = @{
    Id="dc50a0fb-09a3-484d-be87-e023b12c6440"
    Type="Role"
}
Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read Domain.Read.All Directory.ReadWrite.All RoleManagement.ReadWrite.Directory" -DeviceCode
$context = Get-MgContext
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
Write-Host -ForegroundColor Cyan "Certificate loaded"
$appRegistration = New-MgApplication -DisplayName "Leavers_process_OnPrem" -SignInAudience "AzureADMyOrg" -Web @{ RedirectUris="http://localhost"; } -RequiredResourceAccess @{ ResourceAppId=$graphResourceId; ResourceAccess=$UserAuthenticationMethodReadAll,$UserReadWriteAll,$GroupReadWriteAll,$DirectoryReadWriteAll,$ExchangeManageAsApp} -AdditionalProperties @{} -KeyCredentials @(@{ Type="AsymmetricX509Cert"; Usage="Verify"; Key=$cert.RawData })
Write-Host -ForegroundColor Cyan "App registration created with app ID" $appRegistration.AppId
New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
$servicePrincipal = Get-MgServicePrincipal -Filter "displayName eq 'Leavers_process_OnPrem'"
$params = @{
	"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($servicePrincipal.id)"
}
New-MgDirectoryRoleMemberByRef -DirectoryRoleId 0b837ff7-2497-4133-af95-5bac9aa6c423 -BodyParameter $params
Write-Host -ForegroundColor Cyan "Service principal created"
Write-Host
Write-Host -ForegroundColor Green "Success"
Write-Host
$adminConsentUrl = "https://login.microsoftonline.com/" + $context.TenantId + "/adminconsent?client_id=" + $appRegistration.AppId
Write-Host -ForeGroundColor Yellow "Please go to the following URL in your browser to provide admin consent"
Write-Host $adminConsentUrl
Write-Host
Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph"
remove-item $certpath -Force
if ($null -eq $TenantId) {
    $TenantId = (Get-MgOrganization).id
}
$script:dbdata = @{}
$dbdata["OrgName"] = (get-mgdomain | where {$_.id.EndsWith(".onmicrosoft.com")}).id
$dbdata["AppID"] = $appRegistration.AppId
$dbdata['CertThumbprint'] = $cert.Thumbprint
$dbdata['TennantID'] = $TenantId
$dbdata | export-csv $dbstore