
$User = "admin@company.onmicrosoft.com"
$PasswordFile = "Passfile.txt"
$KeyFile = "aes.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -allowclobber

$anon_username = "anonymous"
$anon_password = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
$anon_credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist  $anon_username, $anon_password
$MsgFrom = "TeamsCreationMonitor@company.com"
$smtpserver = "emailserver"
$SmtpPort = '25'
$EmailRecipient = "o365support@company.com"

#set for -7 to do a weekly report, set to -90 to do a quarterly report
$When = ((Get-Date).AddDays(-7)).Date

#set for -7 to do a weekly report, set to -90 to do a quarterly report
$StartDate = (Get-Date).AddDays(-7); $EndDate = (Get-Date)
#HTML header with styles
$htmlhead="
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
    H1{font-size: 22px;}
    H2{font-size: 18px; padding-top: 10px;}
    H3{font-size: 16px; padding-top: 8px;}
    </style>"
#Header for the message
$HtmlBody = "
     <h1>Teams created between $(Get-Date($StartDate) -format g) and $(Get-Date($EndDate) -format g)</h1>
     <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
     <h2><u>Details of Teams Created</u></h2>"
#Person to get the email
# Find records for team creation in the Office 365 audit log
Write-Host "Looking for Team Creation Audit Records, week of $when..."
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "TeamCreated" -ResultSize 1000)
If ($Records.Count -eq 0) {
    Write-Host "No Team Creation records found, week of $when..." }
Else {
    Write-Host "Processing" $Records.Count "audit records..."
    $Report = [System.Collections.Generic.List[Object]]::new()
    ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      $O365Group = (Get-UnifiedGroup -Identity $AuditData.TeamName) # Need some Office 365 Group properties
      $ReportLine = [PSCustomObject]@{
        TimeStamp      = Get-Date($AuditData.CreationTime) -format g
        User           = $AuditData.UserId
        Action         = $AuditData.Operation
        TeamName       = $AuditData.TeamName
        Privacy        = $O365Group.AccessType
        Classification = $O365Group.Classification
        MemberCount    = $O365Group.GroupMemberCount 
        GuestCount     = $O365Group.GroupExternalMemberCount
        ManagedBy      = $O365Group.ManagedBy}
     $Report.Add($ReportLine) }
}
# Add details of each team
$Report | Sort TeamName -Unique | ForEach {
    $htmlHeaderTeam = "<h2>" + $_.TeamName + "</h2>"
    $htmlline1 = "<p>Created on <b>" + $_.TimeStamp + "</b> by: " + $_.User + "</p>"
    $htmlline2 = "<p>Privacy: <b>" + $_.Privacy + "</b> Classification: <b>" + $_.Classification + "</b></p>"
    $htmlline3 = "<p>Member count: <b>" + $_.MemberCount + "</b> Guest members: <b>" + $_.GuestCount + "</b></p>"
    $htmlbody = $htmlbody + $htmlheaderTeam + $htmlline1 + $htmlline2 + $htmlline3 + "<p>"
}
# Finish up the HTML message body    
$HtmlMsg = "" + $HtmlHead + $HtmlBody
# Construct the message parameters and send it off...
 $MsgParam = @{
     To = $EmailRecipient
     From = $MsgFrom
     Subject = "Weekly Teams Creation Report"
     Body = $HtmlMsg
     SmtpServer = $SmtpServer
     Port = $SmtpPort
     Credential = $anon_credentials}
Send-MailMessage @msgParam -BodyAsHTML ; Write-Host "Week of: $when" $EmailRecipient 