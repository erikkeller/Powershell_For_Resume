##############################################################################################################
#
# Global variables
#   Setting these variables is sufficient to run the script, or you can use parameters at the command line.
#   Parameters will override these variables.
# 
$ExpireInDays = 21 # Number of days out to start warning users
$Logging = $true # Set to false to Disable Logging
$_LOGLOCATION = "C:\PasswordNotificationScript\logs\"
$_LOGDATEPATH = Get-Date -Format "MMddyyyy_HHmm"
$_LOGDATETIME = Get-Date
$WhatIf = $false # Set to false to Email Users
$DefaultRecipient = 'itfailednotice@foo.bar' # Default recipient for testing and/or invalid email addresses defined in AD
$AppID = 'registration ID' # App registration ID used to send on behalf of shared mailbox
$DirectoryID = 'tenant ID' # App tenant ID used for connecting to the Graph API
$CertificateThumbprint = 'hash thumbprint' # Thumbprint of Certificate for Graph API Authentication
$BodyTemplate = Get-Content 'C:\PasswordNotificationScript\password_email.html' # Location of the email body template
$SendAccount = 'itnotice@foo.bar' # Shared mailbox to send from
$SearchBase = 'OU=Corp Users,DC=foo,DC=bar' # AD OU to search from
#
##############################################################################################################

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users.Actions

function Export-LogAction
{
    [CmdletBinding()]param
    (
        [Parameter(Mandatory,
        ParameterSetName="Init")]
        [bool]$InitNew,
        [Parameter(Mandatory, 
        ParameterSetName="Log")]$logLevel,
        [Parameter(Mandatory, 
        ParameterSetName="Log")]$logText
    )

    switch ($PSCmdlet.ParameterSetName)
    {
        Init {
            New-Item -Path $($_LOGLOCATION + $_LOGDATEPATH + "_information.txt")
            "   Password expiration notification script Information log file started " + $_LOGDATETIME.ToString() | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_information.txt") -Append

            New-Item -Path $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt")
            "   Password expiration notification script Error log file started " + $_LOGDATETIME.ToString() | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt") -Append
        }
        Log {
            $bodyText = $(Get-Date).ToString() + " - " + $logText

            switch ($logLevel) 
            {
                Information {  
                    $bodyText | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_information.txt") -Append
                }
                Error {
                    $bodyText | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt") -Append
                }
            }
        }
    }
}

# Script quits out if the connection to Graph fails since it's necessary for us to send emails
try 
{
    Connect-MgGraph -ClientID $AppID -TenantId $DirectoryID -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
}
catch 
{
    Write-Host "Error in connecting to Microsoft Graph API"
    Write-Host $_
    exit
}

if ($Logging)
{
    Export-LogAction -InitNew $true
}

$date = Get-Date -format ddMMyyyy

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
Import-Module ActiveDirectory
$ADUsers = Get-ADUser -filter { (Enabled -eq $true) -and (PasswordNeverExpires -eq $false) -and (PasswordExpired -eq $false) }`
    -properties Name,PasswordNeverExpires,PasswordExpired,PasswordLastSet,mail -SearchBase $SearchBase | Where-Object { $_.PasswordLastSet -ne $null }
$MaxAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Process Each User for Password Expiry
foreach ($ADUser in $ADUsers)
{
    $Name = $ADUser.Name
    $EmailAddress = $ADUser.mail
    $PasswordPolicy = (Get-AduserResultantPasswordPolicy $ADUser)
    
    # Check for Fine Grained Password
    if ($null -ne $PasswordPolicy)
    {
        $MaxAge = ($PasswordPolicy).MaxPasswordAge
    }

    $ExpiresOn = ($ADUser.PasswordLastSet) + $MaxAge
    $today = (get-date)
    $ExpirationSpan = (New-TimeSpan -Start $today -End $Expireson).Days

    # Set Greeting based on Number of Days to Expiry.

    if ($ExpirationSpan -gt "1")
    {
        $DaysString = "in " + $ExpirationSpan.ToString() + " days."
    }
    else
    {
        $DaysString = "today."
    }

    # If a user has no email address listed
    if ($null -eq $EmailAddress)
    {
        $emailaddress = $DefaultRecipient
    }
    
    # If a user does not have a foo.bar email address
    elseif (($emailaddress -notlike "*@foo.bar") -and ($emailaddress -notlike "*@baz.bar"))
    {
        $emailaddress = $DefaultRecipient
    }
    # End No Valid Email

    elseif ($WhatIf)
    {
        $EmailAddress = $DefaultRecipient
    }

    # Email Subject Set Here
    $subject="Your password will expire $DaysString"

    $BodyMessage = $BodyTemplate.Replace('$($name)',$name).Replace('$($messageDays)',$DaysString) | Out-String

    # Send Email Message
    #if ($ExpirationSpan -ge "0")
    if (($ExpirationSpan -eq "28") -or ($ExpirationSpan -eq "14") -or ($ExpirationSpan -eq "7") -or ($ExpirationSpan -eq "3") -or ($ExpirationSpan -eq "0"))
    {
        if ($Logging)
        {
            Export-LogAction -logLevel Information -logText "Email sent on $date to $Name ($emailaddress) warning password expires in $ExpirationSpan days."
        }

        $recipients = @()
        $recipients += @{
            emailAddress = @{
                address = $emailaddress
            }
        }

        # Send Email Message
        $message = @{
            subject = $subject;
            toRecipients = $recipients;
            body = @{
                ContentType = 'html';
                Content = $BodyMessage
            }
        }

        Send-MgUserMail -UserId $SendAccount -Message $message

    } # End Send Message
    else # Log Non Expiring Password
    {
        # If Logging is Enabled Log Details
        if ($Logging)
        {
            Export-LogAction -logLevel Information -logText "$Name ($emailaddress) password expires in $ExpirationSpan days."
        }
    }
}
