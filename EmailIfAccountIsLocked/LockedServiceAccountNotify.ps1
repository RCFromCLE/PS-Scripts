#Date deployed: 7/1/2020
#Author: Rudy Corradetti

#email, SMTP technical info
$SMTPHost = "email server"
$SendingEmail = "LockedServiceAccount@company.com"
$Recipients = "ServersGroup@company.com" 
$EmailSubject = "Important Service Account Locked Out, Action Req. - $((Get-Date).ToShortDateString())"
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)



#check if any of these accounts are locked
$systemAccounts = Get-ADGroupMember -identity "LockedServiceAccountNotify-Accounts" -Recursive | Get-ADUser -Property DisplayName, LockedOut |where {$_.LockedOut -eq "True"}|Select name, lockedout
#$systemAccountsFormatted4Email = Get-ADGroupMember -identity "LockedServiceAccountNotify-Accounts" -Recursive | Get-ADUser -Property DisplayName, LockedOut |where {$_.LockedOut -eq "True"}|Select name


#If accounts are locked send an email letting group know they are locked
If ($systemAccounts -like '*True*')
{
$systemAccounts = Get-ADGroupMember -identity "LockedServiceAccountNotify-Accounts" -Recursive | Get-ADUser -Property DisplayName, LockedOut |where {$_.LockedOut -eq "True"}|Select name, lockedout |ConvertTo-Html


#email, build recipients recipients and body
$body1 = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>Attention: Important AD Service Accounts are currently locked out!</h2> </font> 
  
 <font face="Microsoft Tai le"> 
  
The following accounts are locked out. Please unlock the account(s) in AD as fast as possible to avoid any service or processing delays with their associated LOB applications. If the account(s) is not unlocked within 5 minutes, this email will be sent out again.

$systemaccounts
 
</font> 
<br> <br> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by LockedServiceAccountNotify.ps1 running from servername. </i> </font> 
</body> 
"@

#Send email notifying of lockedout account
Send-MailMessage -SmtpServer $SMTPHost -from $SendingEmail -To $Recipients -Subject  "$EmailSubject" -body $body1 -BodyAsHtml -Credential $anonCredentials
}

