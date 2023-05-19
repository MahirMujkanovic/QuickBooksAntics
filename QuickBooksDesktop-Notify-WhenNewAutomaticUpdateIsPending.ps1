#########
# Mahir Mujkanovic @ networkantics.com
# 
#
#########


$senderEmailAddress="mahir.ba@gmail.com"

#Client Secret file path
$clientSecretsDirectory="$env:USERPROFILE\.nasecrets"
$clientSecretFilePath=Join-Path -Path $clientSecretsDirectory -ChildPath "clientSecret.txt"

#Retrieve Client Secret.
if(Test-Path $clientSecretFilePath)
{
    $clientSecretAsSecureString = Get-Content $clientSecretFilePath | ConvertTo-SecureString
    $credential = [pscredential]::new($senderEmailAddress,$clientSecretAsSecureString)
}

# Encrypt and store Client Secret
# Make sure to do this under context of user profile the script will be run under
<######

if(!(Test-Path $clientSecretsDirectory)){New-Item -path $clientSecretsDirectory -ItemType Directory}
$Secure = Read-Host -AsSecureString  #This is where you add the password
$Encrypted = ConvertFrom-SecureString -SecureString $Secure
Set-Content -Value $Encrypted -Path $clientSecretFilePath

######>


######
# FUNCTIONS DECLARATIONS START
######

# Function for sending emails, tested on MS Outlook and/or MS 365 (Exchange online)and GMail email accounts. Use AppPassword if there is MFA enabled
function Send-ToEmail
{
    param(
        [Parameter(Mandatory=$true)][string]$senderEmailAddress,
        [Parameter(Mandatory=$true)][string]$emailTo,
        [Parameter(Mandatory=$false)]$attachmentsPaths,
        [Parameter(Mandatory=$false)]$cred

    )

        $MsgFrom=$senderEmailAddress
        $MsgTo=$emailTo
        #$SmtpServer="smtp.office365.com" ; $SmtpPort="587"
        $SmtpServer="smtp.gmail.com" ; $SmtpPort="587"
        # Build Message Properties
        $Message=New-Object System.Net.Mail.MailMessage $MsgFrom, $MsgTo
        $Message.Subject="<OrgName> - QuickBooks Desktop Automatic Update Pending" 

        #add attachements
        if($attachmentsPaths.Count -gt 0)
        {
            foreach($attachement in $attachmentsPaths)
            {
                $Message.Attachments.Add($attachement)
            }
        }

        $emailBody="QuickBooks Desktop application automatic update is downloaded and pending" + "<br><br>Chop-Chop!! Go and update it!<br><br><br>"
        $Message.Body=$emailBody 
        $Message.IsBodyHTML=$True

        # Force Powershell to use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12

        # Build the SMTP client object and send the message off
        $Smtp=New-Object Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $Smtp.EnableSsl=$True
        $Smtp.credentials=$cred
        $Smtp.Send($Message)
}

# Send-ToEmail -senderEmailAddress $senderEmailAddress -emailTo "hello@mahir.ba" -attachmentsPaths $attachmentsPaths -cred $credential

######
# FUNCTIONS DECLARATIONS END
######


# Check Whether new Update is waiting to be installed

$qbCurrentEditionFilePath="C:\ProgramData\Intuit\QuickBooks 2021\Components\DownloadQB31\Patch\.update\.target\.edition"
$qbLastRememberedEditionFilePath="C:\ProgramData\Intuit\QuickBooks 2021\Components\DownloadQB31\Patch\.update\.target\.edition-lastRemembered"

if(Test-Path -Path $qbLastRememberedEditionFilePath)
{
    $qbLastRememberedEdition=(Get-Content -Path $qbLastRememberedEditionFilePath).Substring(0,10)
    $qbCurrentEdition=(Get-Content -Path $qbCurrentEditionFilePath).Substring(0,10)
} 
else
{
    Copy-Item -Path $qbCurrentEditionFilePath -Destination $qbLastRememberedEditionFilePath
}

# NOTE: Update might automated, investigate on nextt occurence

$qbCurrentEditionFilePath="C:\ProgramData\Intuit\QuickBooks 2021\Components\DownloadQB31\Patch\.update\.target\.edition"
$qbLastRememberedEditionFilePath="C:\ProgramData\Intuit\QuickBooks 2021\Components\DownloadQB31\Patch\.update\.target\.edition-lastRemembered"

if(Test-Path -Path $qbLastRememberedEditionFilePath)
{
    $qbLastRememberedEdition=(Get-Content -Path $qbLastRememberedEditionFilePath).Substring(0,10)
    $qbCurrentEdition=(Get-Content -Path $qbCurrentEditionFilePath).Substring(0,10)
} 
else
{
    Copy-Item -Path $qbCurrentEditionFilePath -Destination $qbLastRememberedEditionFilePath
}

