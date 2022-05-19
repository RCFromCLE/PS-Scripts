##########################################################################
#Start of Connecting to Exchange Online and Security and Compliance Center
##########################################################################
$User = "admin@company.onmicrosoft.com"
$PasswordFile = "Passfile1.txt"
$KeyFile = "aes1.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $MyCredential  -Authentication Basic  -AllowRedirection -ErrorAction Stop
Import-PSSession $Session -allowclobber
##########################################################################
#End of Connecting to Exchange Online and Security and Compliance Center
##########################################################################

$upn = [Environment]::UserName
$errorNumber = Get-Random
$exportlocation = "D:\DownloadedPSTs"
#Using Rudy's location for the Export Tool
$exportexe ="D:\SoftwarePST\DownloadTools\micr..tool_1975b8453054a2b5_000f.0014_34657a1d9b1f16e3\microsoft.office.client.discovery.unifiedexporttool.exe"

#add logging suffix
$suffix = get-date -format MM-dd-yyyy
$dateForLogging = Get-Date -Format "MM/dd/yyyy-HH:mm:sstt"

#So we do not open the queue and interupt the hot file. This will allow see what is querying
$logfileNotFinished = "D:\Logfiles\DownloadPST\NotFinishedMailboxes\MailboxesUnfinished-$suffix.log" 
#Allows us to see what mailboxes were completed that day from the query (Auditing purposes)
$logfileFinished = "D:\Logfiles\DownloadPST\FinishedMailboxes\MailboxesFinished-$suffix.log" 
#If any errors occur in the script
$logfileError = "D:\Logfiles\DownloadPST\DownloadError\DownloaderPSTDownloaderError-$suffix.log" 

#File Path for URL/SAS/Requestor/Email
$filePathForQueue = "D:\Automation\DownloadPST\PSTWaitList\DownloadQueue.csv"

#Count all the lines inside of the file to know how many Mailboxes are in the queue
[int]$LinesInFile = 0
$reader = New-Object IO.StreamReader $filePathForQueue
while ($reader.ReadLine() -ne $null) { $LinesInFile++ } #Open the file and count how many lines are in it
$reader.close() #Closes the file

$LinesInFile -= 1 #Remove one to exclude the Header
########################################################################
#Email Settings 
########################################################################
$anonUsername = "anonymous"
$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anonCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist  $anonUsername, $anonPassword
$SMTPServer = "email server"
$From = "DownloadPST@company.com"
$SubjectError = "PST Error"
$Credential = $anonCredentials
########################################################################
#End of Email Settings
########################################################################

####################################################################
#Start of functions
####################################################################

#This function is responsible for logging Mailboxes that are not finished
Function LogWriteNotFinished { 
    Param ([string]$logstring)
    add-content $logfileNotFinished -value $logstring 
}

#This function is responsible for logging mailboxes that have finished for the day
Function LogWriteFinished { 
    Param ([string]$logstring)
    add-content $logfileFinished -value $logstring 
}

#This function is responsible for catching any errors that might happen
Function LogWriteError { 
    Param ([string]$logstring)
    add-content $logfileError -value $logstring 
}

#This function will email the Requestor that their download has started
function emailUser {
    $body = @"
<font color="Green" face="Microsoft Tai le">
<body BGCOLOR="White">
<h2>Download has started for $userMailbox</h2> </font>


<font face="Microsoft Tai le">

<p>Location of the File: $ExportLocation </p>
<p>Downloads can take up to 24 hours, check within the 24 hours at your own risk</p>
<p>Please login to the Security and Compliance center and remove the associated search.</p>
</font>
<!--mce:0-->

<body BGCOLOR=""white"">

<br> <font face="Microsoft Tai le"> <i>This notification was triggered by DownloadPSTDownloader.ps1 running from companyserver. </i> </font>
</body>
"@

    Send-MailMessage -Body $body -Subject $SubjectComplete  -From $From -To $To -BodyAsHtml -SmtpServer $SMTPServer -Credential $Credential
}

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

<br> <font face="Microsoft Tai le"> <i>This alert was triggered by DownloadPST.ps1 running from companyserver. </i> </font>
</body>
"@

    Send-MailMessage -Body $body -Subject $SubjectError -From $From -To $To -BodyAsHtml -SmtpServer $SMTPServer -Credential $Credential
}

#This will catch if anything were to break
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

####################################################################
#End of functions
####################################################################

########################################################################
#Queue Settings. This information comes from $filePathForQueue. Fills out DL info
########################################################################
$downloadRequest = [PSCustomObject]@{
    UserMailbox   = $searchNameArray #Names the file DL
    URL_Container = $url #Container URL (Need for DL)
    SAS_Key       = $sasKey #SAS Token (Need for DL)
    Requestor     = $request #(Requestor)
}

########################################################################
#End of Queue Settings
########################################################################

#Pull in the download information 
$downloadRequest = Import-Csv -Path $filePathForQueue

#Creates the Temp folder to hold all the nonFinished Mailboxes 
New-Item -Path "D:\Automation\DownloadPST\PSTWaitList" -Name "tempQueue.csv"
$tempQueue = "D:\Automation\DownloadPST\PSTWaitList\tempQueue.csv"
#We have discussed this matter on the headers, but we will leave them stand as they hurt nothing
$changeTemp = Import-Csv $tempQueue -Header UserMailbox, URL_Container, SAS_Key, Requestor 

#Export location of the PST's that are downloaded
$ExportLocation = 'D:\DownloadedPSTs'

#This checks, if the export is fully done. We will then launch the application
for ($i = 0; $i -lt $LinesInFile; $i++) {
    #Builds values to check the information
    $searchname = $downloadRequest[$i].UserMailbox
    $urlSource = $downloadRequest[$i].URL_Container
    $sasKeySource = $downloadRequest[$i].SAS_Key

    #Pull the export information
    $results = Get-ComplianceSearchAction -Identity $searchname | fl results | out-string
    #Search for the progression using regex
    $found = $results -match '.* Progress: (\d+).*'
    
    #if TRUE, it will set the $matches[1] which will equal the progression number
    if ($found) {
        $progression = $matches[1]
    }

    #if $progression equals 100, we will start the download 
    if ($progression -eq 100) {
        $arguments = "-name ""$searchname""", "-source ""$urlSource""", "-key ""$sasKeySource""", "-dest ""$ExportLocation""", "-trace true"
        try {
            Start-Process -FilePath $exportexe -ArgumentList $arguments -ErrorAction Stop
            $mailboxSplit = $downloadRequest[$i].UserMailbox.split("@")
            $userMailbox = $mailboxSplit[0]
            $To = $downloadRequest[$i].Requestor
            Write-Host("Download is beginning for " + $downloadRequest[$i].UserMailbox)
            LogWriteFinished -logstring "##############"
            LogWriteFinished -logstring "Progression is 100%, Download will now begin for $userMailbox"
            LogWriteFinished -logstring "Orginal Requestor: $To"
            $SubjectComplete = "PST Download Started $userMailbox"
            emailUser
        }
        catch {
            exceptionDownloadFailed
        }
    }
    else {
        #This will write to the TempQueue.csv, this will eventually overwrite the DownloadQueue.csv
        $downloadRequest[$i] | Export-Csv -LiteralPath $tempQueue -NoTypeInformation -Append
        LogWriteNotFinished -logstring "######################"
        LogWriteNotFinished -logstring "Date/Time: $dateForLogging"
        LogWriteNotFinished -logstring "Export $searchname has not finished Preparing Data stage"
        LogWriteNotFinished -logstring "Download will reattempt in 10 mins"
        LogWriteNotFinished -logstring "Current progression $progression%"
    }
}

#Delete the old Queue File
remove-item -Path $filePathForQueue
#Make the Temp the Main Queue File
rename-item -Path "D:\Automation\DownloadPST\PSTWaitList\tempQueue.csv" -NewName "DownloadQueue.csv"
while (Get-Process -Name microsoft.office.client.discovery.unifiedexporttool) {
    #This will keep the processing running until it terms
}


#Kill all sessions
Get-PSSession | Remove-PSSession 