#Constants
#
$_CURRENTSTUDENTLISTFILE = "\\fileserver\exchange\ActiveStudents.csv"
$_OFFICE365STUDENTLICENSEGROUP = "CN=Student_O365_Licenses,DC=foo,DC=bar"
$_OFFICE365STUDENTLICENSESKU = "M365EDU_A3_STUUSEBNFT"
$_LOGLOCATION = "C:\O365StudentScript\Log\"
$_LOGDATETIME = Get-Date
$_LOGDATEPATH = $_LOGDATETIME | Get-Date -Format "MMddyyyy_HHmm"
#
#End Constants

#Functions
#
#Add student AD account to the Office 365 license inventory group
function Add-StudentLicenseGroup {
    [CmdletBinding()]param 
    (
        [Parameter(Mandatory)]$ADUser
    )
    
    Write-Output "Adding $($ADUser.sAMAccountName) to Office 365 Student license group"
    Add-ADGroupMember -Identity $_OFFICE365STUDENTLICENSEGROUP -Members $ADUser -Confirm:$false
}
#
#Remove student AD account from the Office 365 license inventory group
function Remove-StudentLicenseGroup {
    [CmdletBinding()]param 
    (
        [Parameter(Mandatory)]$ADUser
    )

    Write-Output "Removing $($ADUser.sAMAccountName) from Office 365 Student license group"
    Remove-ADGroupMember -Identity $_OFFICE365STUDENTLICENSEGROUP -Members $ADUser -Confirm:$false
}
#
#Log script actions to a file
function Export-O365StudentAction
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
            "   Office 365 student license script Information log file started " + $_LOGDATETIME.ToString() | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_information.txt") -Append

            New-Item -Path $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt")
            "   Office 365 student license script Error log file started " + $_LOGDATETIME.ToString() | Out-File -FilePath $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt") -Append
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
#End Functions

#Start the log files
Export-O365StudentAction -InitNew $true

#Connect to Microsoft Graph API
try 
{
    Connect-MgGraph -ClientID <clientID> -TenantID <tenantID> -CertificateThumbprint <certificateThumbprint> -ErrorAction Stop
}
catch
{
    Write-Output "Failed to connect to the Graph API"
    Export-O365StudentAction -logLevel Error -logText "Failed to connect to the Graph API"
    Export-O365StudentAction -logLevel Error -logText "Full error is as follows:"
    Export-O365StudentAction -logLevel Error -logText $_
    exit
}

#Build License SKU object and Servce Plan array
#Service plan array is necessary to remove the Exchange entitlement from Student licenses
$sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $_OFFICE365STUDENTLICENSESKU }
[array]$planToDisable = ($sku.serviceplans | where-object { ($_.ServicePlanName -eq "EXCHANGE_S_ENTERPRISE") -or ($_.ServicePlanName -eq "MCOSTANDARD") }).ServicePlanId

#Grab currently licensed students group
try
{
    $licensedStudents = Get-ADGroup -Identity $_OFFICE365STUDENTLICENSEGROUP -Properties Member -ErrorAction Stop | Select-Object -Expand Member | ForEach-Object { Get-ADUser $_ -Properties mail,sAMAccountName,UserPrincipalName }
}
catch
{
    Write-Output "Failed to retrieve AD accounts from group $_OFFICE365STUDENTLICENSEGROUP"
    Export-O365StudentAction -logLevel Error -logText "Failed to retrieve AD accounts from group $_OFFICE365STUDENTLICENSEGROUP"
    exit
}

#Grab current students
try
{
    $activeStudents = Get-Content -Path $_CURRENTSTUDENTLISTFILE -ErrorAction Stop | Select-Object -Skip 1
}
catch
{
    Write-Output "Failed to retrieve active student list at $_CURRENTSTUDENTLISTFILE"
    Export-O365StudentAction -logLevel Error -logText "Failed to retrieve active student list at $_CURRENTSTUDENTLISTFILE"
    exit
}

#Look for accounts in AD license group but not in active student list
foreach ($licensedStudent in $licensedStudents)
{
    #Somehow a non-student account ended up in the Student License group
    if ($licensedStudent.sAMAccountName -notmatch '\w{2}\d{4,10}')
    { 
        Export-O365StudentAction -logLevel Error -logText "Removing non-student account $($licensedStudent.sAMAccountName) from O365 AD Group"
        Remove-StudentLicenseGroup $licensedStudent
        continue 
    }

    #Check for the account in the list of active students and remove its licenses if it is not and remove it from the Student License group
    if (($activeStudents.Contains("$($licensedStudent.sAMAccountName)@foo.bar")) -or ($activeStudents.Contains("$($licensedStudent.sAMAccountName)@baz.foo.bar")))
    {
        continue
    }
    else 
    {
        $studentAccount = Get-MgUser -UserId $licensedStudent.mail

        if (!$studentAccount)
        {
            $studentAccount = Get-MgUser -UserId "$($licensedStudent.sAMAccountName)@foo.bar"
            
            if (!$studentAccount)
            {
                $studentAccount = Get-MgUser -UserId "$($licensedStudent.sAMAccountName)@baz.foo.bar"

                if (!$studentAccount)
                {
                    Write-Output "Lookup on student $($licensedStudent.sAMAccountName) in O365 License Group failed MgUser Office 365 lookup"
                    Export-O365StudentAction -logLevel Error -logText "Lookup on student $($licensedStudent.sAMAccountName) in O365 License Group failed MgUser Office 365 lookup"
                }
            }
        }

        if ($studentAccount)
        {
            $studentLicenseStatus = Get-MgUserLicenseDetail -UserId $($studentAccount.UserPrincipalName)

            if ($studentLicenseStatus)
            {
                Write-Output "Removing all licenses from $($studentAccount.UserPrincipalName)"
                Export-O365StudentAction -logLevel Information -logText "Removing all licenses from $($studentAccount.UserPrincipalName)"
                $nothing = Set-MgUserLicense -UserId $studentAccount.UserPrincipalName -AddLicenses @() -RemoveLicenses $studentLicenseStatus.SkuId
            }
        }

        Export-O365StudentAction -logLevel Information -logText "Removing student $($licensedStudent.sAMAccountName) from O365 AD Group"
        Remove-StudentLicenseGroup $licensedStudent
    }
}

#Look for students in active student list but do not have an assigned license
foreach ($student in $activeStudents)
{
    #If it's not a student account then skip it
    if ($student -notmatch '\w{2}\d{4,10}@(?:ad\.)?foo\.bar') { continue }

    #Grab associated AD account along with its group membership
    try 
    {
        $studentADAccount = Get-AdUser $($student.Split('@'))[0] -Properties MemberOf -ErrorAction Stop
    }
    catch 
    {
        Write-Output "Student $student not found in AD or other error"
        Export-O365StudentAction -logLevel Error -logText "Student $student not found in AD or other error"
        continue
    }
    
    #Check if account exists in Office 365
    $studentAccount = Get-MgUser -UserId $student

    if (!$studentAccount)
    {
        $adjustedStudent = $student.Split('@')[0] + "@ad." + $student.Split('@')[1]
        $studentAccount = Get-MgUser -UserId $adjustedStudent

        if (!$studentAccount)
        {
            Write-Output "Student $student in active student list but not found in Office 365"
            Export-O365StudentAction -logLevel Error -logText "Student $student in active student list but not found in Office 365"
            continue
        }
    }

    #Check for location set
    $studentLocation = (Get-MgUser -UserId $studentAccount.UserPrincipalName -Property UsageLocation).UsageLocation

    if (!$studentLocation)
    {
        Update-MgUser -UserId $studentAccount.UserPrincipalName -UsageLocation 'US'
    }

    #Check for active license
    $studentLicenseStatus = Get-MgUserLicenseDetail -UserId $studentAccount.UserPrincipalName

    if (!$studentLicenseStatus) 
    {
        #No license assigned
        Write-Output "Adding Office 365 Student license to $($student)"
        Export-O365StudentAction -logLevel Information -logText "Adding Office 365 Student license to $($student)"
        $nothing = Set-MgUserLicense -UserId $student -AddLicenses @{DisabledPlans = $planToDisable;SkuId = ($sku.SkuId)} -RemoveLicenses @()
    }
    elseif ($studentLicenseStatus.SkuPartNumber -notcontains $_OFFICE365STUDENTLICENSESKU) 
    {
        #Incorrect license(s) assigned
        $licensesToRemove = New-Object Collections.Generic.List[string]

        #Ugly hack to get around a bunch of edge cases for pre-existing license assignments
        foreach ($license in $studentLicenseStatus)
        {
            if ($license.SkuPartNumber -eq "STANDARDWOFFPACK_STUDENT")
            {
                $licensesToRemove.Add("<licenseSKU>")
            }
            elseif ($license.SkuPartNumber -eq "STANDARDWOFFPACK_IW_STUDENT")
            {
                $licensesToRemove.Add("<licenseSKU>")
            }
        }

        Write-Output "Adding Office 365 Student license to $student and removing old licenses if they exist"
        Export-O365StudentAction -logLevel Information -logText "Adding Office 365 Student license to $student and removing old licenses if they exist"
        $nothing = Set-MgUserLicense -UserId $student -AddLicenses @{DisabledPlans = $planToDisable;SkuId = ($sku.SkuId)} -RemoveLicenses $licensesToRemove
    }

    if ($studentADAccount.MemberOf -notcontains $_OFFICE365STUDENTLICENSEGROUP)
    {
        Export-O365StudentAction -logLevel Information -logText "Adding student account $($studentADAccount.sAMAccountName) to O365 AD Group."
        Add-StudentLicenseGroup $studentADAccount
    }
}

Disconnect-MgGraph

$resultsBody = "Office 365 Student License Assigment script finished at $((Get-Date).ToString())`n"
$resultsBody += "`n"
$resultsBody += "Accounts with issues are below. For full results, refer to log on ScriptHost at $($_LOGLOCATION + $_LOGDATEPATH + '_information.txt')`n"
$resultsBody += "`n"
$resultsBody += Get-Content $($_LOGLOCATION + $_LOGDATEPATH + "_error.txt") -Raw

Send-MailMessage -From 'O365 Student License Script <scripthost@foo.bar>' -To 'IT Distribution <it_distro@foo.bar>' -Subject "Office 365 Student License Assignment" -Body $resultsBody -SmtpServer smtp.foo.bar