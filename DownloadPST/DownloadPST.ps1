#Purpose: This script is meant to be ran manually to download PST(s) of a user or shared mailbox in Exchange Online, there will be a seperate verison for the SelfServ Portal
#Author: Rudy Corradetti/Matt McConahy 
#Date: 3/17/2022
#Important: Script should be ran on servername, if needing to be moved to another server, AES credentials will need migrated, TFS release process will need updated.

<#
Layout of Script. 
1.) Variables/Email Settings
2.) Functions
3.) Connect to Exchange/Compliance center
4.) Check if mailbox exists 
5.) Search for Mailbox 
6.) Start Export of Mailbox
7.) Start Prepare process
8.) Download Mailbox 
9.) Email user
10.) Term Script

#>

####################################################################
#Start of Variables
####################################################################

$upn = [Environment]::UserName
$exportlocation = "D:\DownloadedPSTs"
$dateForSearch = Get-Date -Format "MMddyyyyHHmmss"
$ServerName = ""
$downloadPath = "D:\Automation\DownloadPST\PSTWaitList\DownloadQueue.csv"
$StillExporting = $true #this will check progress, if the DL is finished, we can term the script
$j = 0 #this will check the DL progress to alert user, it is still working
$i = 0 #this will ccheck the Export progress to alert user, it is still working

$searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
$Requestor = $searcher.FindOne().Properties.mail

########################################################################
#Start of Log Settings 
########################################################################
$suffix = get-date -format MMddyyyy
$logfileError = "D:\Logfiles\DownloadPST\DownloadError\DownloadPSTError-$suffix.txt"
$logfileRan = "D:\Logfiles\DownloadPST\UserRan\UserRan-$suffix.txt"
$dateForLogging = Get-Date -Format "MM/dd/yyyy-HH:mm:ss"

########################################################################
#Log Settings 
########################################################################


########################################################################
#Email Settings 
########################################################################
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist  $anonUsername, $anonPassword
$SMTPServer = "gwmail"
$From = "DownloadPST@company.com"
$To = $requestor
$SubjectError = "PST Error"
$SubjectComplete = "PST Preparing Data"
$Credential = $anonCredentials
########################################################################
#End of Email Settings
########################################################################
#Builds the download information
$downloadRequest = [PSCustomObject]@{
    UserMailbox   = $searchName
    URL_Container = $url
    SAS_Key       = $sasKey
    Requestor     = $requestor
}

####################################################################
#End of Varibles
####################################################################

####################################################################
#Start of functions
####################################################################


#Function will log the current loggged in user running the script
Function LogWriteRan { 
    Param ([string]$logstring)
    add-content $logfileRan -value $logstring 
}
#Function will log errors
Function LogWriteError { 
    Param ([string]$logstring)
    add-content $logfileError -value $logstring 
}
#Function will email user of the Preparing the Data
function emailUser {
    $body = @"
<font color="Green" face="Microsoft Tai le">
<body BGCOLOR="White">
<h2>Preparing data has started for $return</h2> </font>

<font face="Microsoft Tai le">


<p>A script will check the completion of the export 10 minutes</p>
<p>You will be notified when the script has started the download</p>
</font>
<!--mce:0-->

<body BGCOLOR=""white"">

<br> <font face="Microsoft Tai le"> <i>This notification was triggered by DownloadPST.ps1 running from servername. </i> </font>
</body>
"@

    Send-MailMessage -Body $body -Subject $SubjectComplete  -From $From -To $To -BodyAsHtml -SmtpServer $SMTPServer -Credential $Credential
}

#Function will email user of Fatal Error that occured and rec the logs. 
function emailUserError() {

    $body = @"
<font color="Red" face="Microsoft Tai le">
<body BGCOLOR="White">
<h2>Download has failed for $return</h2> </font>

<font face="Microsoft Tai le">


<p>Please logon to $ServerName and review the log D:\Logfiles\DownloadPST\DownloadError.</p>

</font>
<!--mce:0-->

<body BGCOLOR=""white"">

<br> <font face="Microsoft Tai le"> <i>This alert was triggered by DownloadPST.ps1 running from servername. </i> </font>
</body>
"@

    Send-MailMessage -Body $body -Subject $SubjectError -From $From -To $To -BodyAsHtml -SmtpServer $SMTPServer -Credential $Credential
}

#This will throw the error, log that the mailbox was not found (This happens when mailbox does not exist or mistyped)
function exceptionNotFoundMailbox {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "Mailbox was not found $return"
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "MailBox not Found"
    emailUserError
    throw "$Return is not found or unavailable"
    Start-Sleep -Seconds 30
}

#This will throw the error, log that the Search Already Exists,  (This happens when the same request is made twice)
function exceptionSearchExistsAlready {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "Search $SearchName exists in the Compliance Center."
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "Search Exists in Compliance Center"
    emailUserError
    throw "The $SearchName is already in Compliance Center, wait 5 mins to reattempt "
}

#This will throw the error, log that there was an issue with getting the SAS or Token Key. (This happens extremely rare, if there was an issue on microsoft side(Still good to check))
function exceptionGettingSaSorTokenKey {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $requestor"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "SAS key or Token has failed."
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "Getting SAS or Token Key has failed"
    emailUserError
    throw "Unable to get SAS or Token Key "
}

#Testing only exceptions
function exceptionConnectingToOffice365 {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "Connecting to Office 365 has failed"    
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "PST Download is unable to connect to Office 365"
    emailUserError
    throw "Office 365 unable to connect"

    
}
#Testing only exceptions
function exceptionConnectingToComplianceCenter {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "Connecting to Compliance Center has failed"
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "PST Downloader is unable to connect to Compliance Center"
    emailUserError
    throw "Compliance Center unable to connect"

}

#This will throw the error, log that there was issue with the Search. (This happens due to a rare 500 error on microsoft's side.)
function exceptionSearchComplianceError {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "There was an Exception, review error above"
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "There was an issues with Searching/Export"
    emailUserError
    throw "Compliance Search, something went wrong in the search. Possible 500 error "

}

#This will throw the error, log that there was issue with the Export. (This happens if the export failed to grab or the search failed previously)
function exceptionExportComplianceError {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "There was an Exception, review error above"
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    $errorCode = "There was an issue with Searching/Export"
    emailUserError
    throw "Compliance Export, something went wrong in the export. Possible 500 error "

}

function exceptionDownloadFailed {
    LogWriteError -logstring "##############################"
    LogWriteError -logstring "User: $upn"
    LogWriteError -logstring $dateForLogging 
    LogWriteError -logstring $_.Exception.Message
    LogWriteError -logstring "There was an Exception, review error above"
    LogWriteError -logstring "##############################"
    Remove-PSSession -Id $Session.Id #remove from selfService Version
    Remove-PSSession -Id $Session2.Id #remove from selfService Version
    emailUserError
    throw "Download Failed, Check Sas Key and URL Token. Microsoft could've changed location"

}


function button ($title,$mailbx, $WF, $TF) 
{

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName( 'Microsoft.VisualBasic')


$form = New-Object 'System.Windows.Forms.Form';
$form.Width = 500;
$form.Height = 150;
$form.Text = $title;
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

##############Define text label1
$textLabel1 = New-Object 'System.Windows.Forms.Label';
$textLabel1.Left = 25;
$textLabel1.Top = 15;

$textLabel1.Text = $mailbx;


############Define text box1 for input
$textBox1 = New-Object 'System.Windows.Forms.TextBox';
$textBox1.Left = 150;
$textBox1.Top = 10;
$textBox1.width = 200;



#############Define default values for the input boxes
$defaultValue = ''
$textBox1.Text = $defaultValue;


#############define OK button
$button = New-Object 'System.Windows.Forms.Button';
$button.Left = 360;
$button.Top = 85;
$button.Width = 100;
$button.Text = 'OK';

############# This is when you have to close the form after getting values
$eventHandler = [System.EventHandler]{
$textBox1.Text;
$form.Close();};

$button.Add_Click($eventHandler) ;

#############Add controls to all the above objects defined
$form.Controls.Add($button);
$form.Controls.Add($textLabel1);
$form.Controls.Add($textBox1);
$ret = $form.ShowDialog();

#################return values

return $textBox1.Text
}

########################################################################
#End of functions
########################################################################


$return = button "MS365 Email Content Search" "Primary SMTP :"
$datestr = $dateForSearch.ToString()
$SearchName = "$Return-PST-$datestr"

##########################################################################
#Start of Connecting to Exchange Online and Security and Compliance Center
##########################################################################
$User = "admin@company.onmicrosoft.com"
$PasswordFile = "Passfile1.txt"
$KeyFile = "aes1.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $MyCredential  -Authentication Basic  -AllowRedirection -ErrorAction Stop
    Import-PSSession $Session -allowclobber
}
catch {
    exceptionConnectingToComplianceCenter #See function
}
try {
    $Session2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
    Import-PSSession $Session2 -allowclobber
}
catch {
    exceptionConnectingToComplianceCenter #See Function
}
########################################################################
#End of Connecting to Exchange Online and Security and Compliance Center
########################################################################

#Below variables will get the values that had been entered by the user

#Check if the Mailbox Exists
try {
    LogWriteRan -logstring "########################"
    LogWriteRan -logstring "Date/Time: $dateForLogging"
    LogWriteRan -logstring "$upn requested Search/Export for $searchName"

    Write-Host("Checking if mailbox exists")
    Get-Mailbox -Identity $Return -ErrorAction stop
}
catch {
    exceptionNotFoundMailbox #See Function
}

#Check if the Compliance Search has started, if it fails it will throw an exception
try {
    Write-Host("Starting Search Compliance...")
    Start-Sleep -s 3
    New-ComplianceSearch -Name $SearchName -ExchangeLocation $Return -Description "Search all mailbox data for $return"  -AllowNotFoundExchangeLocationsEnabled $true -ErrorAction stop
    Start-ComplianceSearch -Identity $SearchName
}
catch {
    exceptionSearchExistsAlready #See Function
}

########################################################################
#Start of Checking for the Search Completing
########################################################################
do {
    try {
        #Cannot do progress bar here as there is no Progression or number 
        $i += 1
        Start-Sleep -Seconds 30
        $complete1 = Get-ComplianceSearch -Identity $SearchName -ErrorAction stop
        Write-Host("Search is still working, we check every 30 Seconds")
        If ($i -gt 10) {
            Write-Host("Search is still working, searches can take a few minutes ")
            $i = 0
        }
    }
    catch {
        exceptionSearchComplianceError #see Function
    }
}
while ($complete1.Status -ne 'Completed')
Write-Host("Search has completed...")
########################################################################
#End of Checking for the Search Completing
########################################################################


########################################################################
#Start creating Export
########################################################################

#Notify operator of the next process
Write-Host("Starting Export Process...")
Start-Sleep -s 3

# Create Compliance Search in exportable format
New-ComplianceSearchAction -SearchName $SearchName -Export -Format FxStream -ArchiveFormat PerUserPST -EnableDedupe $true -IncludeCredential #| fl
$ExportName = $SearchName + "_Export"

#Wait 30 seconds that the export has a chance to start fully
Write-Host("The script will wait 30 seconds to verify that the export has started")
Start-Sleep -Seconds 30 

try {
    #If we pass the 30 second mark, it is safe to say that the export started successfully
    Get-ComplianceSearchAction -identity $ExportName -ErrorAction stop
    Write-Host("Export is processing")
}
catch {
    exceptionExportComplianceError # This means the export has got a 500 error in Compliance center. This is a microsoft issue, since our mailboxes are in the cloud and not on prem
}
########################################################################
#End of starting Export
########################################################################


######################################################################
#Start of getting Container URL/SAS Token
########################################################################
try {
    #REGEX will be need at some point, but these values are completely random. If it ever breaks debug with VS to check the array to change the index
    Write-Host("We are now grabbing the *Container URL and SAS Token* for the Download")
    Start-Sleep -s 5 
    $index = Get-ComplianceSearchAction -Identity $exportname -includeCredential
    #Split the above values into an array for us to select values
    $y = $index.Results.split(";") 
    #Sets the values to put inside the DownloadQueue.CSV
    $downloadRequest.URL_Container = $y[0].trimStart("Container url: ") #Review $y, if having trouble with download / exectuable portion of script
    $downloadRequest.SAS_Key = $y[1].trimStart(" SAS token: ") #Review $y, if index is incorrect
    $downloadRequest.UserMailbox = $searchName + "_Export"
    $downloadRequest.Requestor = $Requestor[0]
    Write-Host("Values have been grabbed and collected")
    #Writes the Above values inside the DownloadQueue.csv, which is the queue that DownloadPSTDownloader.pst will read in.
    $downloadRequest | export-csv -Path $downloadPath -append -notypeinformation
    Write-Host("Your information has been written to a CSV file, please check back in 24 hours as we are waiting for the Export to Prepare the data. No other actions are required")
    Write-Host("Please check D:\DownloadedPSTs for Download within 24 hours at your own risk...")
    #We are logging who ran DownloadPST, UPN/Date and who you searched for. This will be a day by day item. 
    emailUser
}
catch {
    exceptionGettingSaSorTokenKey #This will very seldom happen, only would happen if the export failed, but if that happens there is another try/catch to mitigate that issue
}
########################################################################
#End of getting Container URL/SAS Token
########################################################################

#Removes Session for Operator Verison of PSTDownloader 
Remove-PSSession -Id $Session.Id
Remove-PSSession -Id $Session2.Id
Start-Sleep -s 30