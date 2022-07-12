##############################################################################################################
#
# Global variables
#   Setting these variables is sufficient to run the script, or you can use parameters at the command line.
#   Parameters will override these variables.
# 
$_ExpireInDays = 21 # Number of days out to start warning users
$_Logging = $true # Set to false to Disable Logging
$_LogFile = "C:\temp\PasswordNotificationLog.csv" # Name and location of the log file
$_WhatIf = $true # Set to false to Email Users
$_DefaultRecipient = 'passwordnotice@foo.bar' # Default recipient for testing and/or invalid email addresses defined in AD
$_AppID = '<insert app registration id>' # App registration ID used to send on behalf of shared mailbox
$_DirectoryID = '<insert app tenant id>' # App tenant ID used for connecting to the Graph API
$_CertificateThumbprint = '<insert certificate thumbprint here>' # Thumbprint of Certificate for Graph API Authentication
$_BodyTemplate = 'C:\temp\EmailNotification.htm' # Location of the email body template
$_SendAccount = 'passwordnotice@foo.bar' # Shared mailbox to send from
$_SearchBase = 'OU=Users,DC=foo,DC=bar' # AD OU to search from
#
##############################################################################################################

function Send-ADUserPasswordNotification
{
    [CmdletBinding()]param
    (
        [string][Parameter()]$AppID = $_AppID,
        [string][Parameter()]$DirectoryID = $_DirectoryID,
        [string][Parameter()]$CertificateThumbprint = $_CertificateThumbprint,
        [int][Parameter()]$ExpireInDays = $_ExpireInDays,
        [string][Parameter()]$SendAccount = $_SendAccount,
        [string][Parameter()]$DefaultRecipient = $_DefaultRecipient,
        [string][Parameter()]$BodyTemplate = $_BodyTemplate,
        [bool][Parameter()]$Logging = $_Logging,
        [string][Parameter()]$LogFile = $_LogFile,
        [bool][Parameter()]$WhatIf = $_WhatIf,
        [string][Parameter()]$SearchBase = $_SearchBase
    )

    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users.Actions

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
        $LogFilePath = Test-Path $LogFile

        if (!$LogFilePath)
        {
            New-Item $Logfile -ItemType File
            Add-Content $Logfile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn,Notified"
        }
    }

    $date = Get-Date -format ddMMyyyy

    # Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
    Import-Module ActiveDirectory
    $ADUsers = Get-ADUser -filter { (Enabled -eq $true) -and (PasswordNeverExpires -eq $false) -and (PasswordExpired -eq $false) }`
        -properties Name,PasswordNeverExpires,PasswordExpired,PasswordLastSet,EmailAddress`
        -SearchBase $SearchBase
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
            continue
        }
        
        # If a user does not have a foo.bar email address
        if (($emailaddress -notlike "*@foo.bar") -and ($emailaddress -notlike "*@baz.bar"))
        {
            #$emailaddress = $testRecipient
            continue
        }
        # End No Valid Email

        # Email Subject Set Here
        $subject="Your password will expire $DaysString"

        $message = $BodyTemplate.Replace('$($name)',$name).Replace('$($messageDays)',$DaysString)

        # Send Email Message
        if (($ExpirationSpan -ge "0") -and ($ExpirationSpan -lt $ExpireInDays))
        {
            if ($Logging)
            {
                Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson,$sent"
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
                    ContentType = "html";
                    Content = $message
                }
            }

            Send-MgUserMail -UserId $SendEmailAccount -Message $message

        } # End Send Message
        else # Log Non Expiring Password
        {
            # If Logging is Enabled Log Details
            if ($Logging)
            {
                Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson,$sent"
            }
        }
    }
}
