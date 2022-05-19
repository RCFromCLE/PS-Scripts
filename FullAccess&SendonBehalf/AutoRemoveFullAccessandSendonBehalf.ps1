#Author: Rudy Corradetti


#Import parameters from SelfServ portal
param
(
[Parameter(Mandatory=$true)][string]$returnSharedMailbox,
[Parameter(Mandatory=$true)][string]$ReturnUserInfo
)

Get-Content -Path $Path

#Build array of users and perform foreach loop to add full and send behalf access for specified shared mailbox
$usrsarray = $ReturnUserInfo.split(",")
foreach ($Returnuser in $usrsarray)

{
    #Remove SoB and FA permissions
    Set-Mailbox "$returnSharedMailbox" -GrantSendOnBehalfTo @{remove= "$ReturnUser"}
    Remove-MailboxPermission -Identity $returnSharedMailbox -User $ReturnUser -AccessRights FullAccess -InheritanceType All -confirm:$false
}