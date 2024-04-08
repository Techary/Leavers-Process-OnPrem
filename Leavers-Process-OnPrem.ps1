param
    (
        [Parameter(Mandatory)][string[]]$upn,
        [switch]$NoPrompt
    )


# ---------------------- USED FUNCTIONS ----------------------
function Get-365Login {
    $dbstoreExists = Test-Path $script:dbstore
    if($dbstoreExists -eq $false) {
        $initialise = read-host "Script has not yet been initialised, do you want to initialise now? (Y/N)"
        if ($initialise -eq "Y") {
            $script:dbdata = @{}
            New-item -path $script:dbstore -ItemType "File"
            $dbdata["username"] = read-host "Enter the UPN of the account that will process the leaver request"
            $dbdata["OrgName"] = read-host "Enter the .onmicrosoft.com domain name"
            $dbdata["appID"] = read-host "Enter the Entra ID enterprise app ID"
            $suffix = ".onmicrosoft.com"
            if (-not $dbdata.OrgName.EndsWith($suffix)) {
                $dbdata.OrgName += $suffix
            }
            $dbdata | export-csv -Path $script:dbstore
            connect-365
        }
        elseif ($initialise -eq "N") {
            write-error "Script cannot be run until the initalisation process has been completed"
            pause
        }
        else {
            write-host "Unkown action, try again"
            start-sleep 5
            get-365login
        }
    }
    else {
        #connect-365
    }
}

function connect-365 {
    function invoke-mfaConnection {
        $dbdata = get-content $script:dbstore   
        Connect-ExchangeOnline -AppId "" `
                                -Certificate "" `
                                -Organization "contoso.onmicrosoft.com" `
                                -ShowBanner:$false
        import-module MSOnline
        import-module ExchangeOnlineManagement
        import-module activedirectory
        Connect-MsolService -Credential $SMTPAuth
    }

    function Get-ExchangeOnlineManagement {

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module -Name ExchangeOnlineManagement

        import-module ExchangeOnlineManagement

        }

    Function Get-MSonline {

        Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted

        Install-Module MSOnline

        }


    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        write-host " "
        write-host "Exchange online Management exists"
    } 
    else {
        Write-host "Exchange Online Management module does not exist. Please ensure powershell is running as admin. Attempting to download..."
        Get-ExchangeOnlineManagement
    }


    if (Get-Module -ListAvailable -Name MSOnline) {
        write-host "MSOnline exists"
    } 
    else {
        Write-host "MSOnline module does not exist. Please ensure powershell is running as admin. Attempting to download..."
        Get-MSOnline
    }

    invoke-mfaConnection

}

function test-upn {

    cls

    if (Get-MsolUser -UserPrincipalName $global:upn -ErrorAction SilentlyContinue)
        {
            Write-host "User found..."
            $global:upn
            $script:upnFound = "True"
        }

    else    
        {
            write-host "$global:upn not found, try again"
            $script:upnFound = "False"
            $Script:UserNotFound = $Script:UserNotFound + $global:upn

        }

}

function disable-localAccount {

    get-aduser -filter "userPrincipalName -eq '$global:upn'" | Disable-ADAccount

}

function get-newpassphrase {

    $SpecialCharacter = @("!","£","`$","%","^","&","*","'","@","~","#")

    $ieObject = New-Object -ComObject 'InternetExplorer.Application'

    $ieObject.Navigate('https://www.worksighted.com/random-passphrase-generator/')

    while ($ieobject.ReadyState -ne 4) {start-sleep -Milliseconds 1}

    $currentDocument = $ieObject.Document

    $password = ($currentDocument.IHTMLDocument3_getElementsByTagName("input") | Where-Object {$_.id -eq "txt"}).value
    $password = $password.Split(' ')[-3..-1]
    $password = -join($password[0],$password[1],$password[2],($SpecialCharacter | Get-Random))

    write-output $password

}   

function set-NewLocalPassword {

    get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADAccountPassword -newpassword $script:SecurePassword

}

function removeLicences {

    $AssignedLicences = (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId

    Invoke-WebRequest -uri https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv -outfile .\licences.csv | Out-Null
    $licences = import-csv .\licences.csv
    remove-item .\licences.csv -Force

    [System.Collections.ArrayList]$UFLicences = @()

    foreach ($Assignedlicence in $Assignedlicences)
        {

            $Assignedlicence = $Assignedlicence.Split(':')[-1]

            foreach ($licence in $licences)
                {

                    if ($Assignedlicence -like $licence."String_ id")
                        {

                            if ($UFLicences -notcontains $licence.Product_Display_name)
                                {

                                    $UFLicences = $UFLicences += $licence.Product_Display_name
                                    
                                }

                        }

                }

        }

    (get-MsolUser -UserPrincipalName $global:upn).licenses.AccountSkuId |
    foreach{
        Set-MsolUserLicense -UserPrincipalName $global:upn -RemoveLicenses $_
    }

}

function Remove-GAL {

        if($script:NoPrompt)
            {

                write-host "Removing from GAL..."
                get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADUser -Replace @{"msDS-CloudExtensionAttribute1"="HideFromGAL"}

                remove-distributionGroups

            }
        else 
            {

                cls

                Write-host "**********************"
                Write-host "** Remove from GAL  **"
                Write-Host "**********************"
                    
                $script:hideFromGAL = Read-Host "Do you want to remove the mailbox from the global address list? ( y / n ) "
                if ($script:hideFromGAL -eq 'Y')
                {
                    get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADUser -Replace @{"msDS-CloudExtensionAttribute1"="HideFromGAL"}

                    Write-host "$global:upn has been hidden"

                    start-sleep 1

                    remove-distributionGroups

                }

                if ($script:hideFromGAL -eq 'N')
                    { 

                        remove-distributionGroups
                        
                    }

                else 
                    {
                        Write-host "You didn't enter an expect response, you idiot."
                        Remove-GAL
                    }
            }
}

function remove-distributionGroups{

    $mailbox = Get-Mailbox -Identity $global:upn
    $DN=$mailbox.DistinguishedName
    $Filter = "Members -like ""$DN"""
    $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter

    if($script:NoPrompt)
        {

            ForEach ($item in $DistributionGroupsList) 
            {

                write-host "Removing from $($item.PrimarySmtpAddress)"
                Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress -Member $global:upn -BypassSecurityGroupManagerCheck -Confirm:$false | Out-Null
                start-sleep 1

            }

        write-result

        }
    else 
        {

            cls

            Write-host "*************************"
            Write-host "** Distribution groups **"
            Write-host "*************************"

            Write-host `n
            Write-host "Listing all Distribution Groups:"
            Write-host `n
            $DistributionGroupsList | ft

            $script:removeDisitri = Read-Host "Do you want to remove $global:upn from all distribution groups ( y / n )?"

            if($script:removeDisitri -eq 'Y')
                {  ForEach ($item in $DistributionGroupsList) 
                    {
                        Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress -Member $global:upn -BypassSecurityGroupManagerCheck -Confirm:$false
                        Write-host "Successfully removed $item"
                        start-sleep 1
                    }
                    Add-Autoreply
                }

            if($script:removeDisitri -eq 'N')
            { Add-Autoreply }

            else {
                write-host "You didn't enter an expect response, you idiot."
                remove-distributionGroups
            }
        }
            
}

function Add-Autoreply {
    cls
        
    Write-Host "***************"
    Write-host "** Autoreply **"
    Write-host "***************"
        
    $script:autoreply = Read-Host "Do you want to add an auto-reply to $global:upn's mailbox? ( y / n / dog ) " 
    if ($script:autoreply -eq 'Y') 
        { $oof = Read-Host "Enter auto-reply"

        Set-MailboxAutoReplyConfiguration -Identity $global:upn -AutoReplyState Enabled -ExternalMessage "$oof" -InternalMessage "$oof"
        write-host "Auto-reply added."
        Add-MailboxPermissions 
        } 

    if ($script:autoreply -eq 'N')      
        { Add-MailboxPermissions } 

    if($script:autoreply -eq 'Dog')
        {   write-host "  __      _"
            write-host  "o'')}____//"
            write-host  " ``_/      )"
            write-host  " (_(_/-(_/"
            start-sleep 5
            Add-Autoreply
        }

    else{ write-host "You didn't enter an expect response, you idiot." 
        Add-Autoreply
        }


}

function Add-MailboxPermissions{
    
    cls
        
    Write-host "*************************"
    Write-host "** Mailbox Permissions **"
    Write-Host "*************************"
        
    $script:mailboxpermissions = Read-Host "Do you want anyone to have access to this mailbox? ( y / n ) "
    if ($script:mailboxpermissions -eq 'y')
        {
            $WhichUser = Read-Host "Enter the E-mail address of the user that should have access to this mailbox "

            add-mailboxpermission -identity $global:upn -user $WhichUser -AccessRights FullAccess

            Write-host "Malibox permisions for $whichUser  have been added"

            write-result

        }

    if($script:mailboxpermissions -eq 'N')
        {
            write-result
        }

                
    else {write-host "You didn't enter an expect response, you idiot."
        Add-MailboxPermissions
            }
}

function write-result {

    if($script:NoPrompt)
        {
            write-host -ForegroundColor red "`n-NoPrompt was used"
            write-host "$global:upn's password has been reset to $script:newlocalpassword"
            write-host "Removed $script:UFLicence"

            pause
        
            continue
        }
    else 
        {
    
            write-host "You have done the following:"

            write-host "`nRemoved $script:UFLicence"

            if ($script:hideFromGAL -eq 'N')
                {
                    write-host -ForegroundColor Yellow "`nYou have not hidden $global:upn from the global address list."
                }
            else
                {
                    write-host -ForegroundColor Green  "`nYou have hidden $global:upn from the global address list."
                }

            if($script:removeDisitri -eq 'N')
                {
                    write-host -ForegroundColor Yellow "`nYou have not removed $global:upn from all distribution groups"
                }
            else
                {
                    write-host -ForegroundColor Green "`nYou have removed $global:upn from any distribution groups."
                }

            if ($script:autoreply -eq 'N')
                {
                    write-host -ForegroundColor Yellow "`nYou have not added an autoreply to $global:upn"
                }
            else 
                {
                    write-host -ForegroundColor Green "`nYou have added an autoreply to $global:upn"
                }

            if($script:mailboxpermissions -eq 'N')
                {
                    write-host -ForegroundColor Yellow "`nYou have not added any mailbox permissions to $global:upn"
                }
            else
                {
                    write-host -ForegroundColor Green "`nYou have added mailbox permissions to $global:upn"
                }

            write-host -ForegroundColor green "`n$global:upn's password has been reset to $script:newlocalpassword"

            pause

            continue
        }

}

function remove-session {

    Write-host "Ending Session..." 

    Disconnect-ExchangeOnline -Confirm:$false | Out-Null

}

function CountDown() {
    param($timeSpan)

    while ($timeSpan -gt 0)
{
    Write-Host '.' -NoNewline
    $timeSpan = $timeSpan - 1
    Start-Sleep -Seconds 1
}
}

# ---------------------- START SCRIPT ----------------------

$today = get-date -format dd-MM-yyyy
Start-Transcript -Append -path "C:\users\$env:username\leaverslog\$today.txt"
$script:NoPrompt = $NoPrompt
$Script:UserNotFound = $null
$script:dbstore = "dbstore.csv"
get-365login

foreach($Global:upn in $upn)
    {
        test-upn

        if($script:upnFound -eq "True")
            {

                disable-localAccount

                $script:NewLocalPassword = Get-Newpassphrase

                $script:SecurePassword = ConvertTo-SecureString $script:NewLocalPassword -AsPlainText -force

                set-NewLocalPassword

                removeLicences

                Set-Mailbox $global:upn -Type Shared

                Remove-GAL
            }

    }
    
if ($null -ne $Script:UserNotFound)
    {

        write-output "`n`nThe following user(s) could not be found: `n$script:userNotFound. `nPlease confirm the email address and try again."

    }

remove-session

Stop-Transcript
