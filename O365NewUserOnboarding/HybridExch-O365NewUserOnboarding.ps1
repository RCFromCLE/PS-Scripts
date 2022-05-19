# Author: Rudy Corradetti
# Change Record:

#If custom exclusions are requested place them here (Up to 6 exclusions at a time, more can be added if needed). Exclusions should not be changed until a new exclusion is needed, the variables must have a value for the script to run
$ExcludeUser1 = ""
$ExcludeUser2 = ""
$ExcludeUser3 = ""
$ExcludeUser4 = ""
$ExcludeUser5 = ""
$ExcludeUser6 = ""



#Check for transcript file, create if not exist then start transcript, if exist start transcript
$Date = (Get-Date).ToString('MM-dd-yyyy')
$TranscriptPath = ""
$TranscriptPathTest = (Test-Path $TranscriptPath)
If ($TranscriptPathTest -eq $false)
    {
    New-Item -ItemType file -Path $TranscriptPath -Force
    Start-Transcript -Path $TranscriptPath -Append
    }
Else
    {
    Start-Transcript -Path $TranscriptPath -Append
    }
#Log successes and errors
$ErrorFile =  ""
$ErrorFileCheck = (Test-Path $ErrorFile)
If ($ErrorFileCheck -eq $false)
    {
    New-Item -ItemType file $ErrorFile
    }
$LogFile =  ""
$LogFileCheck = (Test-Path $LogFile)
If ($LogFileCheck -eq $false)
    {
    New-Item -ItemType file $LogFile -Force
    }

#Connect to Exchange Online using secured credentials
$User = ""
$PasswordFile = "Passfile.txt"
$KeyFile = "aes.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection #|Write-Output "Connecting to Microsoft Exchange Online..."
Import-PSSession $Session -Prefix O365 -allowclobber
#Connect to On premise Exchange
$Session1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer.local/PowerShell/ -Authentication Kerberos
Import-PSSession $Session1 -AllowClobber

#Prepare variables for email, check for errors, send status email
$body1 = @" 
 <font color="Green" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>Successfully onboarded the attached employees to Office 365 and Exchange Online on $Date.</h2> </font> 
  
 <font face="Microsoft Tai le"> 
  
Attached is a log of all users that have been onboarded to Office 365 / Exchange Online on $date.
 
</font> 
<br> <br> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by O365NewUserOnboarding.ps1 running from servername. </i> </font> 
</body> 
"@
$body2 = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>One or more employees failed to be onboarded to Office 365 / Exchange Online on $Date.</h2> </font> 
  
 <font face="Microsoft Tai le"> 

Please logon to servername where automation is running and review O365NewUserLog.txt and Onboarding_Error-$Date to determine which employees failed and repair the failures. 
 
</font> 
<br> <br> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by O365NewUserOnboarding.ps1 running from servername. </i> </font> 
</body> 
"@
$body3 = @" 
 <font color="Red" face="Microsoft Tai le"> 
 <body BGCOLOR="White"> 
<h2>No users onboarded when O365NewUserOnboarding.ps1 ran at $Date!</h2> </font> 
  
 <font face="Microsoft Tai le"> 

When O365NewUserOnboarding.ps1 ran on $Date no new employees were found that needed to be licensed and on boarded to the cloud. If you have received this email more than three days straight please login to CORPADMIN29 and verify the script is running properly.
 
</font> 
<br> <br> 
<!--mce:0--> 
  
<body BGCOLOR=""white""> 
  
<br> <font face="Microsoft Tai le"> <i>This alert was triggered by O365NewUserOnboarding.ps1 running from servername. </i> </font> 
</body> 
"@
$SMTPServer = "Local Exchange Server"
$Recipients = "<RudyCorradetti4@gmil.com>"
$From = "O365Onboarding@company.com"
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)
#Used to generate success log to be emailed to admins
$TempEmailLog = "O365UsersSuccessfullyOnboarded-$Date.txt"
#Used to check if any errors were detected while O365NewUserOnboarding ran
$ErrorCheck = Select-String -Path $ErrorFile -Pattern "Error" -ErrorAction SilentlyContinue


#Grab all users created in last 24 hours who require O365 licensing. For each user set UPN to match primary SMTP and set msExchUsageLocation to US
$When = ((Get-Date).AddDays(-1)).Date
$usrsFormatted4Email = Get-ADUser -Filter {(Enabled -eq $True) -and (WhenCreated -ge $When) -and (department -notlike '') -and (department -notlike '') -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department  -notlike "")  -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (givenname -notlike "") -and (title -ne "") -and (title -ne "") -and (samaccountname -notlike $ExcludeUser1) -and (samaccountname -notlike $ExcludeUser2) -and (samaccountname -notlike $ExcludeUser3)-and (samaccountname -notlike $ExcludeUser4) -and (samaccountname -notlike $ExcludeUser5) -and (samaccountname -notlike $ExcludeUser6)} -Properties department, givenname | Select-Object -ExpandProperty name

#If no users are found needing to be onboarded to O 365 send an email and exit script
If ($usrsFormatted4Email -eq $null)
    {
     Write-Output "No users found needing to be onboarded to Office 365, logging, emailing, and then closing this utility."
     Add-Content -Path $LogFile -Value "$Date - Information: No users found needing to be onboarded to Office 365." -force
     Send-MailMessage -To $Recipients -From $From -Subject "No users found for O365 Onboarding!" -Body $body3 -BodyAsHtml  -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25
     Remove-Item $ErrorFile -Force
     Get-PSSession|Remove-PSSession
     Stop-Transcript
     EXIT
    }
#Grab all users needing to be onboarded, set primary SMTP, set UPN, set location, and licensed status
$usrs = Get-ADUser -Filter {(Enabled -eq $True) -and (WhenCreated -ge $When) -and (department -notlike '') -and (department -notlike '') -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department  -notlike "")  -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (givenname -notlike "") -and (title -ne "") -and (title -ne "") -and -and (samaccountname -notlike $ExcludeUser1) -and  (samaccountname -notlike $ExcludeUser2) -and (samaccountname -notlike $ExcludeUser3)-and (samaccountname -notlike $ExcludeUser4) -and (samaccountname -notlike $ExcludeUser5) -and (samaccountname -notlike $ExcludeUser6)} | Select-Object -ExpandProperty sAMAccountName

#Log users to be onboarded
Add-Content -Path $LogFile -Value "$Date - List of users being onboarded: $usrs" -force


foreach ($usr in $usrs)
{
	#Grab the primary SMTP address
    
    $Company = Get-ADUser -Identity $usr -Properties company | Select -Expand company
    $EmailDomain1 = "@company1.com"
    $EmailDomain2 = "@company2.com"
    $EmailDomain3 = "@company3.com"

    If ($Company -eq "company1")
    {
        Get-AdUser $usr -Properties * |% {Set-ADUser $usr -add @{proxyAddresses="SMTP:"+ $_.GivenName + '.' + $_.Surname +$EmailDomain1}}

        $address = Get-ADUser $usr -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*company1.com*"}
        #Remove the protocol specification from the start of the address
	    $newUPN = $address.SubString(5)

        #Set Location and license
        set-aduser -Identity $usr -UserPrincipalName $newUPN 
        set-aduser -Identity $usr -Replace @{msExchUsageLocation ='US'}
        set-aduser -Identity $usr -Replace @{ExtensionAttribute13 ='Licensed'}

    }

    If ($Company -eq "company2")
    {
        #Add MHS primary SMTP
        Get-AdUser -identity $usr -Properties * |% {Set-ADUser $usr -add @{proxyAddresses="SMTP:"+ $_.GivenName + '.' + $_.Surname +$EmailDomain2}}
        #remove SMTP for medmutual.com
        Get-AdUser -identity $usr -Properties * |% {Set-ADUser $usr -remove @{proxyAddresses="SMTP:"+ $_.GivenName + '.' + $_.Surname +$EmailDomain1}}


        $address = Get-ADUser $usr -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*company2.com*"}
        #Remove the protocol specification from the start of the address
	    $newUPN = $address.SubString(5)

        #Set Location and license
        set-aduser -Identity $usr -UserPrincipalName $newUPN 
        set-aduser -Identity $usr -Replace @{msExchUsageLocation ='US'}
        set-aduser -Identity $usr -Replace @{ExtensionAttribute13 ='Licensed'}

      }
      
      
    If ($Company -eq "company2")
    {
        #Add SDC primary SMTP
        Get-AdUser -identity $usr -Properties * |% {Set-ADUser $usr -add @{proxyAddresses="SMTP:"+ $_.GivenName + '.' + $_.Surname +$EmailDomain3}}
        #remove SMTP for medmutual.com
        Get-AdUser -identity $usr -Properties * |% {Set-ADUser $usr -remove @{proxyAddresses="SMTP:"+ $_.GivenName + '.' + $_.Surname +$EmailDomain}}


        $address = Get-ADUser $usr -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*company3.com*"}
        #Remove the protocol specification from the start of the address
	    $newUPN = $address.SubString(5)

        #Set Location and license
        set-aduser -Identity $usr -UserPrincipalName $newUPN 
        set-aduser -Identity $usr -Replace @{msExchUsageLocation ='US'}
        set-aduser -Identity $usr -Replace @{ExtensionAttribute13 ='Licensed'}

      }    
}
       

#Sync changes to Azure, wait for changes to sync
Try
{ 
Start-ADSyncSyncCycle -PolicyType Delta  
}   
Catch
{
Write-Output "AD Sync currently running, waiting 60 seconds and trying again..."
Start-Sleep -Seconds 60
Start-ADSyncSyncCycle -PolicyType Delta
Start-Sleep -seconds 60
}

#waiting 15 minutes for changes to take effect in O 365
write-output "Waiting 15 minutes for delta sync to complete and cloud to update..."
Start-Sleep -Seconds 900

#Grab users needing to be onboarded to O 365 and confirm they have mailboxes in O 365
$usrs = Get-ADUser -Filter {(Enabled -eq $True) -and (WhenCreated -ge $When) -and (department -notlike '') -and (department -notlike '') -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department  -notlike "")  -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (givenname -notlike "") -and (title -ne "") -and (title -ne "") -and -and (samaccountname -notlike $ExcludeUser1) -and  (samaccountname -notlike $ExcludeUser2) -and (samaccountname -notlike $ExcludeUser3)-and (samaccountname -notlike $ExcludeUser4) -and (samaccountname -notlike $ExcludeUser5) -and (samaccountname -notlike $ExcludeUser6)}| Select-Object -ExpandProperty sAMAccountName
foreach ($usr in $usrs)
{

$address = Get-ADUser $usr -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*"}
$newUPN = $address.SubString(5)


#Begin do until loop to wait until user mailbox has been created, write results to transcript
$CheckIfUserAccountExists = Get-O365mailbox -Identity $newUPN -ErrorAction SilentlyContinue

do 
{
       $CheckIfUserAccountExists = Get-O365mailbox -Identity $newUPN -ErrorAction SilentlyContinue
       Write-Host "Checking if $NewUPN mailbox been created yet"
       Sleep 15
}

Until ($CheckIfUserAccountExists -ne $Null)

Write-Host "$NewUPN mailbox has been created."
   
}
        


#Grab all users created in last 24 hours who require O365 licensing and a functional Exchange Online mailbox
$usrs = Get-ADUser -Filter {(Enabled -eq $True) -and (WhenCreated -ge $When) -and (department -notlike '') -and (department -notlike '') -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department  -notlike "")  -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (department -notlike "") -and (givenname -notlike "") -and (title -ne "") -and (title -ne "") -and -and (samaccountname -notlike $ExcludeUser1) -and  (samaccountname -notlike $ExcludeUser2) -and (samaccountname -notlike $ExcludeUser3)-and (samaccountname -notlike $ExcludeUser4) -and (samaccountname -notlike $ExcludeUser5) -and (samaccountname -notlike $ExcludeUser6)}| Select-Object -ExpandProperty userprincipalname
foreach ($usr in $usrs) 
{

    #Grab each users primary smtp address and nothing else
    $usrSamAccountName = Get-AdUser -Filter {userprincipalname -eq $usr} -Properties samaccountname | Select -ExpandProperty samaccountname
    $email =  Get-ADUser -Filter {SamAccountName -eq $usrSamAccountName} -Properties proxyAddresses | Select -Expand proxyAddresses | Where {$_ -clike "SMTP:*"}
    $formattedemail = $email.SubString(5)

    #Create remote user mailbox in on premise Exchange, set routing address and alias
    Enable-RemoteMailbox $formattedemail -RemoteRoutingAddress $usrsamaccountname@company.mail.onmicrosoft.com
    set-remotemailbox -Identity $formattedemail -Alias $usrsamaccountname

    #Stamp user's Exchange Online mailbox GUID to their corresponding on premise remote mailbox
    $usrExchangeGUID = Get-O365Mailbox -Identity $usr | Format-List ExchangeGUID | Out-String
    $formattedusrExchangeGUID = $usrExchangeGUID.SubString(19)
    Set-RemoteMailbox $usrsamaccountname -ExchangeGuid $formattedusrExchangeGUID

    #Disable users junk mailbox and set retention policy
    Set-O365MailboxJunkEmailConfiguration -Identity $usr -Enabled $false
    Set-O365Mailbox $usr -RetentionPolicy "Default MRM Policy"

    #Make remote user mailbox ACLable
    Set-ADUser -Identity $Usrsamaccountname -Replace @{msExchRecipientDisplayType = -1073741818}
}

write-output "Waiting 15 minutes for cloud users to update..."
#Start-Sleep -Seconds 900



New-Item -ItemType file -path $TempEmailLog -force

foreach ($usr in $usrs) 
 {

$UsrSamAccountName = Get-ADUser -Filter "UserPrincipalName -eq '$usr'" -Properties proxyAddresses | Select-Object -ExpandProperty samaccountname
$RemoteMailboxTest = Get-RemoteMailbox -Identity $usrsamaccountname
$RecDisplayTypeTest = Get-aduser -Identity $usrsamaccountname -Properties msExchRecipientDisplayType | Select-Object -ExpandProperty msExchRecipientDisplayType
$usrExchangeGUIDTest = Get-O365Mailbox -Identity $usr | Format-List ExchangeGUID | Out-String
$formattedusrExchangeGUIDTest = $usrExchangeGUIDTest.SubString(19)
$usrRemoteMailboxExchangeGUIDTest =  get-remotemailbox -Identity $usrsamaccountname | fl ExchangeGUID |Out-String
$FormattedusrRemoteMailboxExchangeGUIDTest = $usrRemoteMailboxExchangeGUIDTest.SubString(19)
$JunkEmailConfigCheck = Get-O365MailboxJunkEmailConfiguration -Identity $usr | select -ExpandProperty Enabled
$RetentionPolicyCheck = Get-O365Mailbox -Identity $usr | select -ExpandProperty Retentionpolicy
get-aduser $usrsamaccountname | select -ExpandProperty userprincipalname | Out-file -FilePath $TempEmailLog -Append -force
    If ($RemoteMailboxTest -eq $null)
        {
        New-Item -ItemType file -Path $ErrorFile -Force -ErrorAction SilentlyContinue
        Add-Content -Path $LogFile -Value "$Date - ERROR: $usr does not have a remote user mailbox in local Exchange. Please attempt to manually create the remote user mailbox and troubleshoot the issue." -force
        Add-Content -Path $ErrorFile -Value "$Date - ERROR: $usr does not have a remote user mailbox in local Exchange. Please attempt to manually create the remote user mailbox and troubleshoot the issue." -force
        }
    
    ELSE{
        Add-Content -Path $LogFile -Value "$Date - Success: $usr does have a remote user mailbox in on premise Exchange." -force
        }

    
    If ($RecDisplayTypeTest -ne -1073741818)
        {
         New-Item -ItemType file -Path $ErrorFile -Force -ErrorAction SilentlyContinue
         Add-Content -Path $LogFile -Value "$Date - ERROR: $usr is not ACLable. Please check the user's msExchRecipientDisplayType and confirm it is set to -1073741818." -force
         Add-Content -Path $ErrorFile -Value "$Date - ERROR: $usr is not ACLable. Please check the user's msExchRecipientDisplayType and confirm it is set to -1073741818." -force
        }
    
    ELSE{
        Add-Content -Path $LogFile -Value "$Date - Success: $usr msExchRecipientDisplayType is ACLable and set properly." -force
        }


    If ($formattedusrExchangeGUIDTest -ne $FormattedusrRemoteMailboxExchangeGUIDTest)
         {
         New-Item -ItemType file -Path $ErrorFile -Force -ErrorAction SilentlyContinue
         Add-Content -Path $LogFile -Value "$Date - ERROR: $usr On Premise Exchange Remote User Mailbox ExchangeGUID does NOT match the Exchange Online Exchange GUID. Please make sure ExchangeGUID from Exchange Online gets set on the remote user mailbox in Exchange On Premise." -force
         Add-Content -Path $ErrorFile -Value "$Date - ERROR: $usr On Premise Exchange Remote User Mailbox ExchangeGUID does NOT match the Exchange Online Exchange GUID. Please make sure ExchangeGUID from Exchange Online gets set on the remote user mailbox in Exchange On Premise." -force
         }

    ELSE{
        Add-Content -Path $LogFile -Value "$Date - Success: $usr OnPremise ExchangeGUID matches their Exchange Online ExchangeGUID." -force
        }

    If ($JunkEmailConfigCheck -eq "True")
        {
        New-Item -ItemType file -Path $ErrorFile -Force -ErrorAction SilentlyContinue
        Add-Content -Path $LogFile -Value "$Date - ERROR: $usr junk email configuration is not disabled, please disable this user's junk email configuration." -Force
        Add-Content -Path $ErrorFile -Value "$Date - ERROR: $usr junk email configuration is not disabled, please disable this user's junk email configuration." -Force
        }

    ELSE
    {
        Add-Content -Path $LogFile -Value "$Date - Success: $usr junk email configuration is disabled."
    }

    If ($RetentionPolicyCheck -ne "Default MRM Policy")
    {
    New-Item -ItemType file -Path $ErrorFile -Force -ErrorAction SilentlyContinue
    Add-Content -Path $LogFile -Value "$Date - ERROR: $usr retention policy is not set to Default MRM Policy, please set the retention policy properly." -Force
    Add-Content -Path $ErrorFile -Value "$Date - ERROR: $usr retention policy is not set to Default MRM Policy, please set the retention policy properly." -Force
    }

    ELSE
    {
    Add-Content -Path $LogFile -Value "$Date - Success: $usr retention policy is set to Default MRM Policy."
    }    

 }

#If error file contains the string error send an email letting admin know there is an error else send a success email with attached users who were onboarded successfully
if ($ErrorCheck -ne $null)
    {
     Write-Output "Errors were detected during O365 onboarding!"
     Send-MailMessage -To $Recipients -From $From -Body $body2 -BodyAsHtml -Subject "O365 Onboarding Failed!" -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25
    }
    ELSE
    {
     Send-MailMessage -To $Recipients -From $From -Body $body1 -BodyAsHtml -Subject "O365 Onboarding Succeeded!" -Attachments $TempEmailLog -SmtpServer $SMTPServer -Credential $anonCredentials -Port 25
     Remove-Item $TempEmailLog -Force
     Remove-Item $ErrorFile -Force
    }

Get-PSSession|Remove-PSsession
Stop-Transcript



      



    