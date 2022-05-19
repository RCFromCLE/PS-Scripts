#Author: Rudy Corradetti

#Import parameters from SelfServ portal
param
(
[Parameter(Mandatory=$true)][string]$returnSharedMailbox,
[Parameter(Mandatory=$true)][string]$ReturnUserInfo
)

#Build array of users and perform foreach loop to add full and send behalf access for specified shared mailbox
$usrsarray = $ReturnUserInfo.split(",")
foreach ($Returnuser in $usrsarray)

{
    #Configure SoB and FA permissions
    Set-Mailbox "$returnSharedMailbox" -GrantSendOnBehalfTo @{add= "$Returnuser"}
    Add-MailboxPermission -Identity "$returnSharedMailbox" -User "$Returnuser" -AccessRights FullAccess -InheritanceType All
}