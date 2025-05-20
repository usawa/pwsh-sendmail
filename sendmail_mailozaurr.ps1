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
.PARAMETER TestConn
    If specified, the script will check if the SMTP:Port is accessible by opening a socket.
.PARAMETER V
    Verbose mode, it will display some debug logs
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
    [Parameter(Mandatory)][String]$SMTPServer,
    [Parameter(Mandatory)][Int32]$Port,
    [Parameter(Mandatory)][String]$From,
    [Parameter(Mandatory)][String]$To,
    [Parameter(Mandatory)][String]$Subject,
    [Parameter(Mandatory)][String]$Body,
    [switch]$UseSSL,
    [String]$CC,
    [String]$BCC,
    [String]$User,
    [String]$Passwd,
    [String]$AttachmentPath,
    [Switch]$High,
    [switch]$TestConn,
    [switch]$V,
    [switch]$DryRun
)

#Installed with:  Install-Module -Name Mailozaurr -AllowPrerelease
# Needs Powershell 5.1+
Import-Module Mailozaurr

# Send-EmailMessage [-Server <String>] [-Port <Int32>] [-From <Object>] [-ReplyTo <String>] [-Cc <String[]>]
# [-Bcc <String[]>] [-To <String[]>] [-Subject <String>] [-Priority <String>] [-Encoding <String>]
# [-DeliveryNotificationOption <String[]>] [-DeliveryStatusNotificationType <DeliveryStatusNotificationType>]
# [-Credential <PSCredential>] [-SecureSocketOptions <SecureSocketOptions>] [-UseSsl] [-HTML <String[]>]
# [-Text <String[]>] [-Attachment <String[]>] [-Timeout <Int32>] [-Suppress] [-WhatIf] [-Confirm]
# [<CommonParameters>]

# Global value
$global:SMTPPorts = @( 25, 465, 587 )

# Valid Email Address
function IsValidEmail {
    param([string]$EmailAddress)
    $regex = '^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$'

    return [regex]::IsMatch($EmailAddress, $regex)
}

# Valid FQDN
function IsValidFQDN {
    param([String]$FQDN)
    $regex = '^(?!:\/\/)(?=.{1,255}$)((.{1,63}\.){1,127}(?![0-9]*$)[a-z0-9-]+\.?)$'

    return [regex]::IsMatch($FQDN, $regex)
}

# Valid IPv4
function IsValidIPv4 {
    param([String]$IPAddress)
    try {
        [System.Net.IPAddress]::Parse($IPAddress) | Out-Null
        return $true
    } catch {
        return $false
    }
}

function IsValidSMTPPort {
    param(
        [int]$Port,
        [switch]$Enforce = $false
    )

    if ($Port -lt 1 -or $Port -gt 65535) {
        return $false;
    }
    if($Enforce) {
        if (!($global:SMTPPorts -contains $Port)) {
            return $false;
        }
    }
    return $true
}

# Function to test port connectivity, using sockets
function Test-Port
{
    param ( 
        [string]$Hostname, 
        [int]$Port, 
        [int]$Milliseconds = 500 
    )
 
    #  Initialize object
    $Test = New-Object -TypeName Net.Sockets.TcpClient
 
    #  Attempt connection, 300 millisecond timeout, returns boolean
    ( $Test.BeginConnect( $Hostname, $Port, $Null, $Null ) ).AsyncWaitHandle.WaitOne( $Milliseconds )
 
    # Cleanup
    $Test.Close()
}


# Params Hashlist
$Parameters = @{}

# Check SMTP Server valid or not
if ( !( (IsValidFQDN -FQDN $SMTPServer) -or (IsValidIPv4 -IPAddress $SMTPServer) ) ) {
    Write-Error "SMTP Server: $SMTPServer isn't a valid FQDN or IP address"
    exit 1
}

$Parameters.Add("Server", $SMTPServer)

# Check if Port is valid
if ( !(IsValidSMTPPort -Port $Port) ) {
    Write-Error "SMTP Port: $Port isn't a valid value"
    exit 1
}
$Parameters.Add("Port", $Port)

# Now test port
if ($TestConn -eq $true) {
    if ( !(Test-Port -Hostname $SMTPServer -Port $Port) ) {
        Write-Error "SMTP Server $SMTPServer`:$Port isn't accessible. Please check your firewall rules"
        exit 1
    } else {
        if ($V) {
            Write-Output "$SMTPServer`:$Port is opened"
        }
    }
}

# Priority
if ($High) {
    $Parameters.Add("Priority", "High")
}

#SSL
if ($UseSSL) {
    $Parameters.Add("UseSsl","")
}

# Subject
$Parameters.Add("Subject", $Subject)

# Use secure connection if available
$UseSecureConnectionIfAvailable = $true
$Parameters.Add("UseSecureConnectionIfAvailable", $UseSecureConnectionIfAvailable)

# SMTP: Sender
if (IsValidEmail -EmailAddress $From) {
    $SMTPSender = [MimeKit.MailboxAddress]$From
    $Parameters.Add("From", $SMTPSender)
} else {
    Write-Error "From: $From isn't a valid Email Address"
    exit 1
}

# SMTP: To
$RecipientArray = $To -split ","
#$SMTPRecipientList = [MimeKit.InternetAddressList]::new()
foreach($Recipient in $RecipientArray) {
    if( !(IsValidEmail -EmailAddress $Recipient) ) {
        Write-Error "To: $Recipient isn't a valid email address"
        exit 1
    }
}

$Parameters.Add("To", $To)

# SMTP: CC
if ($CC) {
    $CCArray = $CC -split ","
    foreach($Recipient in $CCArray) {
        if( !(IsValidEmail -EmailAddress $Recipient) ) {
            Write-Error "CC: $Recipient isn't a valid email address"
            exit 1
        }
    }
    $Parameters.Add("Cc", $CC)
}

# SMTP: BCC
if ($BCC) {
    $BCCArray = $BCC -split ","
    foreach($Recipient in $BCCArray) {
        if(! (IsValidEmail -EmailAddress $Recipient) ) {
            Write-Error "BCC: $Recipient isn't a valid email address"
            exit 1
        }
    }
    $Parameters.Add("Bcc", $Bcc)
}

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
#Let's convert it (will add a parameter)
$Body = $Body.replace("\n","`n")
$Parameters.Add("Text", $Body)

# HTML Body, not managed yet. We'll see if this is needed
#$HTMLBody = [string]"HTMLBody"
#$Parameters.Add("HTMLBody", $HTMLBody)

# Attachement files, only one here, test if exists
if($AttachmentPath) {
    $AttachmentArray = $AttachmentPath -split ","
    $AttachmentList = [System.Collections.Generic.List[string]]::new()
    foreach($Attachment in $AttachmentArray) {
        if(Test-Path -path $Attachment) {
            $AttachmentList.Add($Attachment)
        } else {
            Write-Error "Attachment file: $AttachmentPath can't be found"
            exit 1
        }
    }
    $Parameters.Add("Attachments", $AttachmentPath)
}

if ($v) {
    Write-Output $Parameters
    Write-Output "Recipients: $RecipientArray"
    Write-Output "CC: $CCArray"
    Write-Output "BCC: $BCCArray"
}

# Send Email
if ( !$DryRun ) {
    Send-MailMessage @Parameters
}