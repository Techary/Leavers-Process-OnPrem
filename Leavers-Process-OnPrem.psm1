function invoke-leaverprocess {
    #requires -module exchangeonlinemanagement
    #requires -module activedirectory
    #requires -module microsoft.graph

    param (
            [Parameter(Mandatory)][string[]]$upn,
            [switch]$NoPrompt
    )
    function Get-365Login {
        $dbstoreExists = Test-Path $script:dbstore
        if($dbstoreExists -eq $false) {
            write-error "Setup script has not been run. Please (re)run the setup script, and try again."
        }
        else {
            connect-365
        }
    }
    function connect-365 {
        import-module activedirectory
        import-module exchangeonlinemanagement
        import-module microsoft.graph
        $dbdata = import-csv $script:dbstore
        Connect-ExchangeOnline  `
            -AppId $dbdata.appid `
            -CertificateThumbprint $dbdata.CertThumbprint `
            -Organization $dbdata.orgname `
            -ShowBanner:$false
        Connect-MgGraph `
            -ClientId $dbdata.appid `
            -TenantId $dbdata.TennantID `
            -CertificateThumbprint $dbdata.CertThumbprint `
            -ContextScope Process `
            -NoWelcome
    }
    function test-upn {
        $script:userObject = get-mguser -filter "UserPrincipalName eq '$upn'"
        if ($script:userObject) {
            Write-host "`nUser found!"
            $script:upnFound = "True"
            start-sleep 1
        }
        else {
            write-host "User not found"
            $script:upnFound = "False"
            $Script:UserNotFound = $Script:UserNotFound + $global:upn
            start-sleep 1
            get-upn
        }
    }
    function disable-localAccount {
        get-aduser -filter "userPrincipalName -eq '$global:upn'" | Disable-ADAccount
    }
    function get-stringhash {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $input
        )
        if ($null -eq $input ) {
            $input = read-host "Enter string"
        }
        $stringAsStream = [System.IO.MemoryStream]::new()
        $writer = [System.IO.StreamWriter]::new($stringAsStream)
        $writer.write($input)
        $writer.Flush()
        $stringAsStream.Position = 0
        Get-FileHash -InputStream $stringAsStream | Select-Object Hash
    }
    function set-NewLocalPassword {
        get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADAccountPassword -newpassword $script:SecurePassword
    }
    function removeLicences {
        $AssignedLicences = (get-mguserlicenseDetail -userid $script:userObject.id)
        $ProgressPreference = 'silentlycontinue'
        Invoke-WebRequest -uri https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv -outfile .\licences.csv | Out-Null
        $licences = import-csv .\licences.csv
        remove-item .\licences.csv -Force
        [System.Collections.ArrayList]$script:UFLicences = @()
        foreach ($Assignedlicence in $Assignedlicences) {
            foreach ($licence in $licences) {
                if ($Assignedlicence.skupartnumber -like $licence.String_id) {
                    if($script:UFLicences -notcontains $licence.Product_Display_name) {
                        $script:UFLicences.add($licence.Product_Display_name) | out-null
                    }
                }
            }
        }
        if ($script:UFLicences.count -eq 0) {
            clear-host
            write-host -ForegroundColor red "There are no licences applied to this account."
            $continue = read-host "Do you want to contine? YOU WILL SEE ERRORS. Y/N"
            switch ($continue) {
                Y {
                    $script:NoLicence = 'y'
                }
                N {
                    write-result
                }
                default {
                    removeLicences
                }
            }
        }
        else {
            try {
                Set-MgUserLicense -userID $script:userObject.id -AddLicenses @() -RemoveLicenses @($Assignedlicences.skuid)
            }
            catch {
                $_.exception[0]
                $script:LicenceRemovalError = $true
            }
        }
    }
    function Remove-GAL {
        if($script:NoPrompt) {
            write-host "Removing from GAL..."
            get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADUser -Replace @{"msDS-CloudExtensionAttribute1"="HideFromGAL"}
            remove-distributionGroups
        }
        else  {
            cls
            Write-host "**********************"
            Write-host "** Remove from GAL  **"
            Write-Host "**********************"  
            $script:hideFromGAL = Read-Host "Do you want to remove the mailbox from the global address list? ( y / n ) "
            if ($script:hideFromGAL -eq 'Y') {
                get-aduser -filter "userPrincipalName -eq '$global:upn'" | Set-ADUser -Replace @{"msDS-CloudExtensionAttribute1"="HideFromGAL"}
                Write-host "$global:upn has been hidden"
                start-sleep 1
                remove-distributionGroups
            }
            if ($script:hideFromGAL -eq 'N') { 
                remove-distributionGroups   
            }
            else  {
                Write-host "You didn't enter an expect response, you idiot."
                Remove-GAL
            }
        }
    }
    function remove-distributionGroups {
        if($script:NoPrompt) {
            ForEach ($item in $DistributionGroupsList)  {
                write-host "Removing from $($item.PrimarySmtpAddress)"
                Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress -Member $global:upn -BypassSecurityGroupManagerCheck -Confirm:$false | Out-Null
                start-sleep 1
                write-result
            }
        }
        else {
            Clear-Host
            print-TecharyLogo
            Write-host "*************************"
            Write-host "** Distribution groups **"
            Write-host "*************************"
            $mailbox = Get-Mailbox -Identity $script:userobject.userprincipalname
            $DN=$mailbox.DistinguishedName
            $Filter = "Members -like ""$DN"""
            $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
            Write-host `n
            Write-host "Listing all Distribution Groups:"
            Write-host `n
            $DistributionGroupsList | ft
            Do {
                $script:removeDisitri = Read-Host "Do you want to remove $($script:userobject.userprincipalname) from all distribution groups ( y / n )?"
                Switch ($script:removeDisitri) {
                    Y {
                        ForEach ($item in $DistributionGroupsList) {
                            $RemovalException = $false
                            try {
                                #Not migrated to MGGraph as there's no native powershell function to manipulate groups
                                Remove-DistributionGroupMember -Identity $item.PrimarySmtpAddress -Member $script:userobject.userprincipalname -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction stop
                            }
                            catch {
                                Write-Output "Unable to remove from $($item.displayname)"
                                $_.exception[0]
                                $script:RemovalException = $true
                            }
                            finally {
                                if($RemovalException -eq $false) {
                                    Write-host "Successfully removed from $($item.DisplayName)"
                                }
                            }
                        }
                        Add-Autoreply
                    }
                    Default {
                        "You didn't enter an expect response, you idiot."
                    }
                    N {
                        Add-Autoreply
                    }
                }

            }
            until ($script:removeDisitri -eq 'y' -or $script:removeDisitri -eq 'n')
        }
    }
    function Add-Autoreply {
        Do {
            Clear-Host
            print-TecharyLogo
            Write-Host "***************"
            Write-host "** Autoreply **"
            Write-host "***************"
            $script:autoreply = Read-Host "Do you want to add an auto-reply to $($script:userobject.userprincipalname)'s mailbox? Y will add a new OOF; N will turn off any pre-existing OOF; Leave won't make any changes but, will show you a good boy. ( y / n / leave ) "
            Switch ($script:autoreply) {
                Y {
                    $oof = Read-Host "Enter auto-reply"
                    try {
                        #Not supported in MGGraph yet
                        set-MailboxAutoReplyConfiguration -Identity $script:userobject.userprincipalname -AutoReplyState Enabled -ExternalMessage "$oof" -InternalMessage "$oof" -ErrorAction stop
                    }
                    catch {
                        Write-output "Unable to set auto-reply"
                        $_.exception
                        $script:AutoReplyError = $true
                    }
                    finally {
                        if($null -eq $AutoReplyError) {
                            write-host "Auto-reply added."
                        }
                    }
                    Add-MailboxPermissions
                }
                N {
                    Set-MailboxAutoReplyConfiguration -Identity $script:userobject.userprincipalname -AutoReplyState Disabled
                }
                Default {
                    "You didn't enter an expect response, you idiot."
                }
                Leave {
                    write-host "  __      _"
                    write-host  "o'')}____//"
                    write-host  " `_/      )"
                    write-host  " (_(_/-(_/"
                    start-sleep 5
                    Add-MailboxPermissions
                }
            }

        }
        until ($script:autoreply -eq 'y' -or $script:autoreply -eq 'n' -or $script:autoreply -eq 'Dog')
    }
    function Add-MailboxPermissions {
        Do {
            Clear-Host
            print-TecharyLogo
            Write-host "*************************"
            Write-host "** Mailbox Permissions **"
            Write-Host "*************************"
            $script:mailboxpermissions = Read-Host "Do you want anyone to have access to this mailbox? ( y / n ) "
            Switch ($script:mailboxpermissions) {
                Y {
                    $script:WhichUserPermissions = Read-Host "Enter the E-mail address of the user that should have access to this mailbox "
                    if(get-mguser -filter "UserPrincipalName eq '$script:WhichUserPermissions'" -ErrorAction SilentlyContinue) {
                        try {
                            add-mailboxpermission -identity $script:userobject.userprincipalname -user $script:WhichUserPermissions -AccessRights FullAccess -erroraction stop
                        }
                        catch {
                            write-output "Unable to add permissions"
                            $_.exception[0]
                            $script:MailboxError = $true
                        }
                        finally {
                            if($null -eq $MailboxError) {
                                Write-host "Malibox permisions for $script:WhichUserPermissions have been added"
                            }
                        }
                        }
                    else {
                        Write-output "$script:WhichUserPermissions not found. Please try again"
                        start-sleep 3
                        Add-MailboxPermissions
                    }
                    Add-MailboxForwarding
                }
                N {
                    Add-MailboxForwarding
                }
                Default {
                    "You didn't enter an expect response, you idiot."
                }
            }
        }
        until ($script:mailboxpermissions -eq 'y' -or $script:mailboxpermissions -eq 'n')
    }
    function Add-MailboxForwarding {
        Do {
            Clear-Host
            print-TecharyLogo
            Write-host "*************************"
            Write-host "** Mailbox Forwarding **"
            Write-Host "*************************"
            $script:mailboxForwarding = Read-Host "Do you want any forwarding in place on this account? ( y / n ) "
            Switch ($script:mailboxForwarding) {
                Y { 
                    $script:WhichUserForwarding = Read-Host "Enter the E-mail address of the user that emails should be forwarded to "
                    if(get-mguser -filter "UserPrincipalName eq '$script:WhichUserForwarding'" -ErrorAction SilentlyContinue) {
                        try {
                            Set-Mailbox $script:userobject.userprincipalname -ForwardingAddress $script:WhichUserForwarding -erroraction stop
                        }
                        catch {
                            write-output "Unable to add permissions"
                            $_.exception[0]
                            $script:ForwardingError = $true
                        }
                        finally {
                            if($null -eq $script:ForwardingError) {
                                Write-host "Malibox forwarding to $script:WhichUserForwarding has been added"
                            }
                        }
                    }
                    else {
                        Write-output "$script:WhichUserForwarding not found. Please try again"
                        start-sleep 3
                        Add-MailboxForwarding
                    }
                    write-result
                    }
                N {
                    write-result
                }
                Default {
                    "You didn't enter an expect response, you idiot."
                }
            }
        }
        until ($script:mailboxforwarding -eq 'y' -or $script:mailboxforwarding -eq 'n')
    }
    function write-result {
        Clear-Host
        write-host "You have done the following:"
        switch ($script:LicenceRemovalError) {
            $true {
                write-host -ForegroundColor Red "`nThere was an error attempting to removing the licences from this account. Please review the log $psscriptroot\logs\$($script:userobject.userprincipalname).txt"
            }
            default {
                switch ($script:NoLicence) {
                    Y {
                        write-host -ForegroundColor yellow "No licences were assigned to this account."
                    }
                    default {
                        write-host "`nRemoved the following licence(s):" ; $script:UFLicences
                    }
                }
            }
        }
        switch ($script:GALError){
            $true {write-host -ForegroundColor Red "`nThere was an error hiding from the GAL. Please review the log $psscriptroot\logs\$($script:userobject.userprincipalname).txt"}
            default {
                switch ($script:hideFromGAL) {
                        N {
                            write-host -ForegroundColor Yellow "`nYou have not hidden $($script:userobject.userprincipalname) from the global address list."
                        }
                        Y {
                            write-host -ForegroundColor Green  "`nYou have hidden $($script:userobject.userprincipalname) from the global address list."
                        }
                }
            }
        }
        switch ($script:RemovalException) {
            $true {
                write-host -ForegroundColor Red "`nThere was an error removing $($script:userobject.userprincipalname) from some distribution lists. Please review the log $psscriptroot\logs\$($script:userobject.userprincipalname).txt"
            }
            default {
                switch ($script:removeDisitri) {
                    Y {
                        write-host -ForegroundColor Green "`nYou have removed $($script:userobject.userprincipalname) from all distribution groups."
                    }
                    N {
                        write-host -ForegroundColor Yellow "`nYou have not removed $($script:userobject.userprincipalname) from all distribution groups"
                    }
                }
            }
        }
        switch ($script:AutoReplyError) {
            $true {
                Write-host -ForegroundColor red "`nThere was an error adding the auto reply. Plese review the log $psscriptroot\logs\$($script:userobject.userprincipalname).txt"
            }
            default {
                switch ($script:autoreply) {
                    N {
                        write-host -ForegroundColor Yellow "`nYou have not added an autoreply to $($script:userobject.userprincipalname)"
                    }
                    Y {
                        write-host -ForegroundColor Green "`nYou have added an autoreply to $($script:userobject.userprincipalname)"
                    }
                }
            }   
        }
        switch ($script:MailboxError) {
            $true {
                Write-Host -ForegroundColor red "`nThere was an error adding the mailbox permissions. Please review the log $psscriptroot\$($script:userobject.userprincipalname)"
            }
            default {
                switch ($script:mailboxpermissions) {
                    N {
                        write-host -ForegroundColor Yellow "`nYou have not added any mailbox permissions to $($script:userobject.userprincipalname)"
                    }
                    Y {
                        write-host -ForegroundColor Green "`nYou have added mailbox permissions for $script:whichuserPermissions to $($script:userobject.userprincipalname)"
                    }
                }
            }
        }
        switch ($script:ForwardingError) {
            $true {
                write-host -ForegroundColor red "`nThere was an error adding the email forwarding. Please review the log $psscriptroot\logs\$($script:userobject.userprincipalname).txt"
            }
            default {
                switch ($script:mailboxForwarding) {
                    N {
                        write-host -ForegroundColor Yellow "`nYou have not added any mailbox forwarding to $($script:userobject.userprincipalname)"
                    }
                    Y {
                        write-host -ForegroundColor Green "`nYou have added mailbox forwarding to $script:WhichUserForwarding"
                    }
                }
            }
        }
        switch ($script:refreshTokenError) {
            $true {
                write-host -ForegroundColor red "`nFailed to revoke the refresh tokens. Any current active sessions will remain active until autentication token expires"
            }
            default {}
        }
        switch ($script:SetPassswordError ) {
            $true {
                write-host -ForegroundColor red "`nThere was an error setting the password on this account. Please check the log at $psscriptroot\logs\$($script:userobject.userprincipalname).txt"
            }
            default {
                write-host -ForegroundColor green "`nSet password to $($script:NewCloudPassword.password)"
            }
        }
        switch ($script:ConversionFailure) {
            $true {
                write-host -ForegroundColor red "`nThere was an error converting the mailbox to shared. Please see the log in $psscriptroot\$($script:userobject.userprincipalname).txt"
            }
            Default {}
        }
        Write-Host "`nA transcript of all the actions taken in this script can be found at $psscriptroot\$($script:userobject.userprincipalname).txt"
        pause
    }
    function remove-session {
        Write-host "Ending Session..." 
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Disconnect-MgGraph -ErrorAction SilentlyContinue | out-null
    }
    function CountDown() {
        param($timeSpan)
        while ($timeSpan -gt 0) {
        Write-Host '.' -NoNewline
        $timeSpan = $timeSpan - 1
        Start-Sleep -Seconds 1
        }
    }
    # ---------------------- START SCRIPT ----------------------
    $today = get-date -format dd-MM-yyyy
    Start-Transcript -Append -path "C:\users\$env:username\leaverslog\$today.txt"
    $Script:UserNotFound = $null
    $script:dbstore = "dbstore.csv"
    get-365login
    foreach($Global:upn in $upn) {
        test-upn
        if($script:upnFound -eq "True") {
            disable-localAccount
            $script:NewLocalPassword = ((get-stringhash $upn).hash + (get-random) + "%")
            $script:SecurePassword = ConvertTo-SecureString $script:NewLocalPassword -AsPlainText -force
            set-NewLocalPassword
            removeLicences
            Set-Mailbox $global:upn -Type Shared
            Remove-GAL
        }
    }
    if ($null -ne $Script:UserNotFound) {
        write-output "`n`nThe following user(s) could not be found: `n$script:userNotFound. `nPlease confirm the email address and try again."
        }
    remove-session
    Stop-Transcript
}
