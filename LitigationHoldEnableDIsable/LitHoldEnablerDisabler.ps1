#Author: Rudy Corradetti


#Use Encrypted credentials to login as PSadmin to Exchange Online 
$User = "admin@company.onmicrosoft.com"
$PasswordFile = "Passfile.txt"
$KeyFile = "aes.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection #|Write-Output "Connecting to Microsoft Exchange Online..."
Import-PSSession $Session -allowclobber

#Import AD Module, Get all members of App-LitigationHold, select only LANID
Import-Module ActiveDirectory
$membername = Get-ADGroupMember –Identity “LitigationHoldADGroup” |Select -expand SamAccountName

#Create file for email, Get all user mailboxes where SamAccountName is in APP-LitigationHold and LitigationHold is not enabled and enable it
$Outfile = "litholdusers1.txt"
if ( -Not (Test-Path "$outfile"))
{
New-Item -path $Outfile -ItemType file -force -ErrorAction Continue
}
Clear-Content $Outfile
$ListWithoutLitHold = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited | where{ ($_.Alias -in $membername) –and ($_.LitigationHoldEnabled -match "False") } |fl Name | Tee-Object -FilePath $Outfile
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited | where{ ($_.Alias -in $membername) –and ($_.LitigationHoldEnabled -match "False") } | Set-Mailbox -LitigationHoldEnabled $True

#Declare SMTP variables, configure email formatting
$SMTPHost = "emailserver"
$SendingEmail = "LitHold-Automation@company.com"
$Recipients = "<Rudy.Corradetti@company.com>"
[string[]]$To = $Recipients.Split(',')
$EmailSubject = "Lit-Hold Automation Report - $((Get-Date).ToShortDateString())"
$MessageBody = $ListWithoutLitHold | out-string
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

#Determine if sending an email is needed, delete or clear content of litholdusers.txt
If((Get-Content $Outfile) -ne $null)
{
Send-MailMessage -SmtpServer $SMTPHost -from $SendingEmail -To $To -Subject  "$EmailSubject" -body "Litigation Hold has been enabled for the following users - $MessageBody" -Credential $anonCredentials
}

#Create file for email, Get all user mailboxes where SamAccountName is not in APP-LitigationHold and LitigationHold is enabled and disable it
$Outfile2 = "C:\ExchangeOnlineAutomation\PS Scripts\litholdusers2.txt"
if ( -Not (Test-Path "$outfile2"))
{
New-Item -path $Outfile2 -ItemType file -force -ErrorAction Continue
}
Clear-Content $Outfile2
$ListWithoutLitHold2 = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited | where{ ($_.Alias -notin $membername) –and ($_.LitigationHoldEnabled -match "True") }  |fl Name | Tee-Object -FilePath $Outfile2
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited | where{ ($_.Alias -notin $membername) –and ($_.LitigationHoldEnabled -match "True") } | Set-Mailbox -LitigationHoldEnabled $False

#Declare SMTP variables, configure email formatting
$SMTPHost2 = "emailserver"
$SendingEmail2 = "LitHold-Automation@company.com"
$Recipients2 = "<Rudy.Corradetti@company.com>"
[string[]]$To2 = $Recipients2.Split(',')
$EmailSubject2 = "Lit-Hold Automation Report - $((Get-Date).ToShortDateString())"
$MessageBody2 = $ListWithoutLitHold2 | out-string

#Determine if sending an email is needed, delete or clear content of litholdusers.txt
If((Get-Content $Outfile2) -ne $null)
{
Send-MailMessage -SmtpServer $SMTPHost2 -from $SendingEmail2 -To $To2 -Subject  "$EmailSubject2" -body "Litigation Hold has been disabled for the following users - $MessageBody2" -Credential $anonCredentials
Get-PSSession|Remove-PSSession
Exit
}

If((Get-Content $Outfile2) -eq $null)
{

    If ((Get-Content $Outfile) -eq $null) 
    {
    Send-MailMessage -SmtpServer $SMTPHost -from $SendingEmail -To $To -Subject $EmailSubject2 -Body "There have been no changes to the LitigationHold AD group." -Credential $anonCredentials 
    Remove-Item $Outfile
    Remove-Item $Outfile2
    Get-PSSession|Remove-PSSession
    EXIT
    }
}

#Remove-Item $Outfile
#Remove-Item $Outfile2


#Close all open PS sessions
Get-PSSession|Remove-PSSession