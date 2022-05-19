#Set to true to enable, set to false to disable
$Active = "true"

#Configure amount of messages needing to be in the mailbox queue before email alerts are sent out
$threshold = "500"

#Connect to Exchange
$Session1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://wdcpexch01.mms.ams.local/PowerShell/ -Authentication Kerberos
Import-PSSession $Session1 -AllowClobber


#Variables needed for email
$Date = "{0:h:mm:ss tt zzz}" -f (get-date) -replace ".{7}$"
$SMTPServer = "Gwmail.mmoh.com"
$From = "ExchangeQueueMonitor@medmutual.com"
$Recipients = "<ExchangeSuppport@medmutual.com>", "<Todd.Ryan@medmutual.com>"
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)


#Get message queue size, email if above threshold for WDCPEXCH01
$results = Get-Queue -Server "WDCPEXCH01" | Select Identity, MessageCount
foreach($result in $results)
{
  [string]$ServerName = $result.Identity  
  [int]$MessageCount = $result.MessageCount
}

  IF ($MessageCount -gt $threshold -and $Active -eq "true") 
  {
    #MessageCount is greater than 500
    $body = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>WDCPEXCH01 has greater than $threshold messages in queue.</h2> </font> 
  
 <font face="Microsoft Tai le"> 

<p>Total Mesages in Queue ($date) : $MessageCount </p>

<p>Please logon to $ServerName and review why the mail queue is greater than $threshold.</p> 

<p>This email will continue to be sent every five minutes until the queue is lower then $threshold.</p>
 
</font> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by MailQueueMonitor.ps1 running from WDCPADMIN02. </i> </font> 
</body> 
"@

   Send-MailMessage -To $Recipients -From $From -Subject "Mail Queue on $ServerName has greater than $threshold messages" -Body $body -BodyAsHtml  -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25

  }



#Get message queue size, email if above threshold for WDCPEXCH02
$results = Get-Queue -Server "WDCPEXCH02" | Select Identity, MessageCount
foreach($result in $results)
{
  [string]$ServerName = $result.Identity 
  [int]$MessageCount = $result.MessageCount
}

  IF ($MessageCount -gt $threshold -and $Active -eq "true") 
  {
    #MessageCount is greater than 500
    $body = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>WDCPEXCH02 has greater than $threshold messages in queue.</h2> </font> 
  
 <font face="Microsoft Tai le"> 

<p>Total Mesages in Queue ($date) : $MessageCount </p>

<p>Please logon to $ServerName and review why the mail queue is greater than $threshold.</p> 

<p>This email will continue to be sent every five minutes until the queue is lower then $threshold.</p>

 
</font> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by MailQueueMonitor.ps1 running from WDCPADMIN02. </i> </font> 
</body> 
"@

   Send-MailMessage -To $Recipients -From $From -Subject "Mail Queue on $ServerName has greater than $threshold messages" -Body $body -BodyAsHtml  -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25

  }



#Get message queue size, email if above threshold for WDCPEXCH03
$results = Get-Queue -Server "WDCPEXCH03" | Select Identity, MessageCount
foreach($result in $results)
{
  [string]$ServerName = $result.Identity
  [int]$MessageCount = $result.MessageCount
}

  IF ($MessageCount -gt $threshold -and $Active -eq "true") 
  {
    #MessageCount is greater than 500
    $body = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>WDCPEXCH03 has greater than $threshold messages in queue.</h2> </font> 
  
 <font face="Microsoft Tai le"> 

<p>Total Mesages in Queue ($date) : $MessageCount </p>

<p>Please logon to $ServerName and review why the mail queue is greater than $threshold.</p> 

<p>This email will continue to be sent every five minutes until the queue is lower then $threshold.</p>
 
</font> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by MailQueueMonitor.ps1 running from WDCPADMIN02. </i> </font> 
</body> 
"@

   Send-MailMessage -To $Recipients -From $From -Subject "Mail Queue on $ServerName has greater than $threshold messages" -Body $body -BodyAsHtml  -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25

  }






$Date = "{0:h:mm:ss tt zzz}" -f (get-date)
$formatteddate -replace ".{6}$"
