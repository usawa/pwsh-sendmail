<#
.SYNOPSIS
    A Powerhsell script to send emails using SMTP and Send-MailkitMessage module
.DESCRIPTION
    This script takes some parameters to send an email using SMTP server. As based 
    on MailKit and MimeKit, it supports SMPT with or without Authentication, STARTTLS with 
    latest TLS versions (depending of TLS supported version by Windows).
    It accepts multiple Recipients (TO), and a file attachment.
.PARAMETER SMTPServer
    The name or IP Address of the SMTP server.
.PARAMETER Port
    The port of SMTP server. There's no default so you must specify, for example, 25 or 587.
.PARAMETER From
    The email address of the sender. There's a syntax check based on regular expression.
.PARAMETER To
    One or more recipients, comma separated (a@a.com,b@b.com, ...). 
    There's a syntax check based on regular expression.
.PARAMETER Subject
    Subject of the email
.PARAMETER Body
    The body of the message, simple text (not HTML). Note the carriage return isn't \n 
    but `n with Powershell
.PARAMETER User
    If SMTP server requires authentication, it's the username or login.
    If password isn't set, authentication is ignored.
.PARAMETER Passwd
    If SMTP server requires authentication, it's the password. 
    If user isn't set, authentication is ignored.
.PARAMETER AttachmentPath
    The System Path to an Attachment file (file sent with mail)
    If file isn't here it will fail.
.PARAMETER V
    Verbose mode, it will display some debug logs
.PARAMETER Help
    Please run Get-Help sendmail.ps1 -Detailed.
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Sébastien Rohaut
    Date:   May 16, 2025 
.LINK
    https://github.com/austineric/Send-MailKitMessage/tree/master
.LINK
    https://github.com/usawa
#>

param(
    [Parameter(Mandatory)]
    [string]$SMTPServer,
    [Parameter(Mandatory)]
    [int]$Port,
    [Parameter(Mandatory)]
    [string]$From,
    [Parameter(Mandatory)]
    [string]$To,
    [Parameter(Mandatory)]
    [String]$Subject,
    [Parameter(Mandatory)]
    [String]$Body,
    [String]$User,
    [String]$Passwd,
    [String]$AttachmentPath,
    [switch]$V,
    [switch]$Help
)

Import-Module Send-MailKitMessage

# Valid Email Address
function IsValidEmail {
    param([string]$EmailAddress)
    $regex = '^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$'

    return [regex]::IsMatch($EmailAddress, $regex)
}


# Display Help
if($Help -eq $true) {
    Write-Host "Help!"
    exit 0
}

# Params Hashlist
$Parameters = @{}

$Parameters.Add("SMTPServer", $SMTPServer)
$Parameters.Add("Port", $Port)
$Parameters.Add("Subject", $Subject)

# Use secure connection if available
$UseSecureConnectionIfAvailable = $true
$Parameters.Add("UseSecureConnectionIfAvailable", $UseSecureConnectionIfAvailable)

# SMTP: Sender
if(IsValidEmail($From)) {
    $SMTPSender = [MimeKit.MailboxAddress]$From
    $Parameters.Add("From", $SMTPSender)
} else {
    Write-Error "From: $From isn't a valid Email Address"
    exit 1
}

# SMTP: To
$RecipientArray = $To -split ","
$SMTPRecipientList = [MimeKit.InternetAddressList]::new()
foreach($Recipient in $RecipientArray) {
    if(IsValidEmail($Recipient)) {
        $SMTPRecipientList.Add([MimeKit.InternetAddress]$Recipient)
    } else {
        Write-Error "To: $Recipient isn't a valid email address"
        exit 1
    }
}
$Parameters.Add("RecipientList", $SMTPRecipientList)


# SMTP: Creds, not mandatory
if ($User -and $Passwd) {
    if($V) {
        Write-Output "Using Credentials with user: $User"
    }
    $Crendential = [System.Management.Automation.PSCredential]::new($User, (ConvertTo-SecureString -String $Passwd -AsPlainText -Force))
    $Parameters.Add("Credential", $Crendential)
}

# text body
# Note that to send newlines, it's not \n but `n (powershell)
$Parameters.Add("TextBody", $Body)

# HTML Body, not managed yet. We'll see if this is needed
#$HTMLBody = [string]"HTMLBody"
#$Parameters.Add("HTMLBody", $HTMLBody)

# Attachement files, only one here, test if exists
if($AttachmentPath) {
    if(Test-Path -path $AttachmentPath) {
        $AttachmentList = [System.Collections.Generic.List[string]]::new()
        $AttachmentList.Add($AttachmentPath)
        $Parameters.Add("AttachmentList", $AttachmentList)
    } else {
        Write-Error "Attachment file: $AttachmentPath can't be found"
        exit 1
    }
}

if ($v) {
    Write-Output $Parameters
    Write-Output "Recipients: $RecipientArray"
}

# Send Email
Send-MailKitMessage @Parameters