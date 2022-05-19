######Author - Rudy Corradetti

# See Intel doc for username and password info on psadmin
$username = "admin@company.onmicrosoft.com"
$password = "#####"
$secureStringPwd = ConvertTo-SecureString $password -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd

$secureStringText = $secureStringPwd | ConvertFrom-SecureString  
Set-Content "D:\Automation\Encrypted Passwords\Passfile1.txt" $secureStringText 



# Generate a random AES Encryption Key.
$AESKeyFilePath = "aes1.key"
$AESKey = New-Object Byte[] 16
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
	
# Store the AESKey into a file. This file should be protected!  (e.g. ACL on the file to allow only select people to read)
Set-Content $AESKeyFilePath $AESKey   # Any existing AES Key file will be overwritten		

$password = $secureStringPwd | ConvertFrom-SecureString -Key $AESKey
Add-Content $secureStringPwd $password





$KeyFile = "aes1.key"
$Key = New-Object Byte[] 32 # You can use 16, 24, or 32 for AES
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
$Key | out-file $KeyFile
$PasswordFile = "Passfile1.txt"
$Key = Get-Content $KeyFile
$Password = "#####" | ConvertTo-SecureString -AsPlainText -Force
$Password | ConvertFrom-SecureString -key $Key | Out-File $PasswordFile