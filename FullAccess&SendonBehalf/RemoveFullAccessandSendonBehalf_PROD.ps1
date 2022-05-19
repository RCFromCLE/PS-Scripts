#Use Encrypted credentials to login as 
$User = "admin@company.onmicrosoft.com"
$PasswordFile = "Passfile.txt"
$KeyFile = "aes.key"
$key = Get-Content $KeyFile
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential ` -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $MyCredential -Authentication Basic -AllowRedirection #|Write-Output "Connecting to Microsoft Exchange Online..."
Import-PSSession $Session -allowclobber


#If admin chooses Y continue running script over and over
$choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Y","&N")
while ( $true ) {
  
#Get Shared Mailbox alias and user LAN ID
function button ($title,$mailbx, $WF, $TF) 

{

[void][System.Reflection.Assembly]::LoadWithPartialName( “System.Windows.Forms”)
[void][System.Reflection.Assembly]::LoadWithPartialName( “Microsoft.VisualBasic”)


$form = New-Object “System.Windows.Forms.Form”;
$form.Width = 500;
$form.Height = 150;
$form.Text = $title;
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

#Define text label1
$textLabel1 = New-Object “System.Windows.Forms.Label”;
$textLabel1.Left = 25;
$textLabel1.Top = 15;

$textLabel1.Text = $mailbx;

#Define text box1 for input
$textBox1 = New-Object “System.Windows.Forms.TextBox”;
$textBox1.Left = 150;
$textBox1.Top = 10;
$textBox1.width = 200;

#Define default values for the input boxes
$defaultValue = “”
$textBox1.Text = $defaultValue;

#define OK button
$button = New-Object “System.Windows.Forms.Button”;
$button.Left = 360;
$button.Top = 85;
$button.Width = 100;
$button.Text = “Ok”;

# This is when you have to close the form after getting values
$eventHandler = [System.EventHandler]{
$textBox1.Text;
$form.Close();};

$button.Add_Click($eventHandler) ;

#Add controls to all the above objects defined
$form.Controls.Add($button);
$form.Controls.Add($textLabel1);
$form.Controls.Add($textBox1);
$ret = $form.ShowDialog();

#return values
return $textBox1.Text
}

#Below variables will get the values that have been entered by the admin
$returnSharedMailbox = button “Shared Mailbox alias or email:”
$returnUserInfo = button “User LAN ID:”

$SharedMailbox = Get-Mailbox -Identity $ReturnSharedMailbox -ErrorAction SilentlyContinue 
$UserMailbox = Get-Mailbox -Identity $returnUserInfo -ErrorAction SilentlyContinue

#Evaluate values entered by admin
If ($SharedMailbox -ne $null -and $UserMailbox -ne $null)   

{
    #Set Path to currently logged in user's desktop
    $DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)

    #Remove SoB and FA permissions
    Set-Mailbox "$returnSharedMailbox" -GrantSendOnBehalfTo @{remove= "$returnUserInfo"}
    Remove-MailboxPermission -Identity $returnSharedMailbox -User $returnUserInfo -AccessRights FullAccess -InheritanceType All -confirm:$false

    #Audit SoB and FA permissions, store in variable, create text file, add send on behalf header
    $SharedMBEmail = $SharedMailbox.PrimarySmtpAddress
    Write-output "Below is the list of users with Send on Behalf to $SharedMBEmail -" `n | Out-file -FilePath "$DesktopPath\FullAccess&SendonBehalfCheck.txt"
    Get-Mailbox "$returnSharedMailbox" | Select -ExpandProperty GrantSendOnBehalfTo | Out-File -FilePath "$DesktopPath\FullAccess&SendonBehalfCheck.txt" -Append -Force
    Write-output `n `n "Below is the List of users with Full Access to $SharedMBEmail -" | Out-File "$DesktopPath\FullAccess&SendonBehalfCheck.txt" -Append -Force
    Get-MailboxPermission $returnSharedMailbox | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") } |FL Identity, user, AccessRights | Out-File -FilePath "$DesktopPath\FullAccess&SendonBehalfCheck.txt" -Append -Force

    #Pop up information, store email address in variable
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("You have removed full access and send on behalf access for $returnUserInfo on $SharedMBEmail. Please check your email for a complete listing of send on behalf and full access permissions on $SharedMBEmail.")
    
    #Send email to currently logged in user's email address containing success and attachment with all SoB and FA info for $SharedMailbox
    $Attachment = "$DesktopPath\FullAccess&SendonBehalfCheck.txt"
    $LoggedinRecipient = ([adsi]"LDAP://$(whoami /fqdn)").mail   
    $recipient = ([adsi]"LDAP://$(whoami /fqdn)").mail
    $body = “You have removed full access and send on behalf for $UserMailbox on $SharedMBemail! See attachment for complete list of send on behalf and full access permissions on $SharedMBEmail.”
    Send-MailMessage -To $LoggedinRecipient -from "$LoggedinRecipient" -Subject "Removed FA & SoB for $UserMailbox on $SharedMBEmail" -Body $body -BodyAsHtml -Attachments $Attachment -smtpserver emailserver -Port 25
    
    #Delete $Attachment from desktop
    Remove-Item -Path $Attachment
}ELSE
{
    #Notify admin that they entered an incorrect value
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("User Mailbox or Shared Mailbox not found in Exchange Online! Confirm the primary values entered are correct and each mailbox exists in the cloud.")
}
  #Ask to run again 
  $choice = $Host.UI.PromptForChoice("Run Again?","",$choices,0)
  if ( $choice -ne 0 ) {
    break
  }
}

Get-PSSession|Remove-PSSession
