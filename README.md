# Leavers-Process-CloudOnly
*THIS IS ONLY FOR 365 WITH SYNC TO ON PREMISE DIRECTORY*

### Converts a leavers mailbox to shared, removes the licence, asks if you want to:  
1. Remove from GAL  
2. Remove from disitribution lists  
3. Add an auto reply  
4. Add read+manage permissions  
5. Add mailbox forwarding

### Prerequisites
[Git](https://git-scm.com/downloads) must be installed

[Powershell 7+](https://github.com/PowerShell/PowerShell/releases/tag/v7.4.1) should be installed

A custom rule must be set in ADConnect that looks for msDS-CloudExtensionAttribute1 to be set to HideFromGAL (https://www.uclabs.blog/2023/06/how-to-hide-users-from-gal-if-they-are.html)

### How to use
#### Installation
1. Ensure [Git](https://git-scm.com/downloads) is installed
2. Ensure [Powershell 7+](https://github.com/PowerShell/PowerShell/releases/tag/v7.4.1) is installed. (Built in PS7+ but 5+ will _probably_ work)
3. `cd` into `C:\users\$env:username\documents\powershell\modules`
4. Run `git clone https://github.com/Techary/Leavers-Process-OnPrem.git`
5. `cd` into the newly created folder
6. Run `.\setup.ps1`.
7. Sign in with an account with access to these scopes:
   
   Application.ReadWrite.All
   
   User.Read

   Domain.Read.All
   
   Directory.ReadWrite.All
   
   RoleManagement.ReadWrite.Directory

8. Accept the admin request
9. Follow the instructions on the CLI (it should give you a link to follow to grant further admin consent)   
10. Run `import-module Leavers_Process-OnPrem.ps1`
#### Usage
1. Open powershell as an administrator
2. Run `invoke-leaversprocess <upn>`
