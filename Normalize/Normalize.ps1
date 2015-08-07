<# Normalize.ps1

Normalizes an onsite computer to the new Citrix, RightFax (prompted), and SSO program.

    08032015 - Creation - EK

#>

# Check for admin rights
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

$Rightfax = $false

While ($true)
{
    Clear-Host
    Write-Host "**************************************************"
    Write-Host "*                                                *"
    Write-Host "*           Normalize Windows 7 PC               *"
    Write-Host "*                                                *"
    Write-Host "* Normalize a PC for new Citrix, Single Sign-On, *"
    Write-Host "* and Rightfax, if necessary                     *"
    Write-Host "*                                                *"
    Write-Host "*   !!Computer will force restart at the end!!   *"
    Write-Host "**************************************************"
    Write-Host "Install Rightfax? (Y/N/Q to quit):" -NoNewline
    $x = [Console]::ReadKey()
    $x = [char]::ToUpper($x.KeyChar)
    if ($x -eq "Y") { $Rightfax = $true; Break }
    elseif ($x -eq "Q") { Exit }
    elseif ($x -eq "N") { Break }
    else
    {
        Write-Host "`nInvalid keypress, please use either Y, N, or Q"
        Write-Host "Press any key to continue ..."
        $y = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

Clear-Host
Write-Host "Checking/Installing Prereqs ..."
#Check Pre-Reqs
# 378389 is the min version of .NET 4.0, and this key wouldn't exist 
# if .NET 4.x wasn't installed, and since NULL is less than 378389 
# it passes the logic anyway
if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Net Framework Setup\NDP\v4\Full").Release -lt 378389) 
{
    Start-Process "\\SoftwareRepo\RightFax\Prereqs\dotnetfx45_full_x86_x64.exe" -ArgumentList "/q /norestart" -Verb RunAs -Wait
}
#All the checks for VC++ redists were terrible (random registry keys 
# for each version? no install dir? 3 different product codes for each version? 
# Thanks Microsoft) while reinstalls are short and non-invasive so we'll do it the easy way
Start-Process "\\SoftwareRepo\RightFax\Prereqs\VS2008_vcredist_x86.exe" -ArgumentList "/q /norestart" -Verb RunAs -Wait
Start-Process "\\SoftwareRepo\RightFax\Prereqs\VS2010_vcredist_x86.exe" -ArgumentList "/q /norestart" -Verb RunAs -Wait
Start-Process "\\SoftwareRepo\RightFax\Prereqs\VS2012_vcredist_x86.exe" -ArgumentList "/q /norestart" -Verb RunAs -Wait

Write-Host "Checking for and uninstalling old version of Vergence ..."
# Check for old version of Vergence and uninstall
$applications = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
    Get-ItemProperty |
    where { $_.Publisher -eq "Sentillion, Inc." }
if ($applications -ne $null)
{
    Get-Process lp, trayProxy, authenticator, C2W_CM | Stop-Process -Force # wow this is easy
    Get-Service VergenceSessionManager, VergenceLocatorSvc | Stop-Service -Force # still easy
    Remove-Item -Path "$env:ProgramFiles\Sentillion" -Force -Recurse
    Remove-Item -Path "$env:ProgramData\Sentillion" -Force -Recurse
    foreach ($app in $applications)
    {
        Start-Process "msiexec" -ArgumentList "/X$($app.PSChildName) /qb /norestart /passive" -Wait -Verb RunAs 
    }
}
else { Write-Host "   Skipped" }

Write-Host "Checking for and uninstalling old version of Citrix ..."
# Uninstall old version of Citrix, if it exists
$applications = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
    Get-ItemProperty |
    where { $_.DisplayName -eq "Citrix Presentation Server Client" }
if ($applications -ne $null)
{
    foreach ($app in $applications)
    {
        Start-Process "msiexec" -ArgumentList "/X$($app.PSChildName) /qb /norestart /passive" -Wait -Verb RunAs 
    }
}
else { Write-Host "   Skipped" }

Write-Host "Installing new Citrix if necessary ..."
# Check to see if new Citrix is not already installed
if (!(Test-Path "C:\Program Files\Citrix\SelfServicePlugin\SelfServicePlugin.exe"))
{
	# Install new version of Citrix
	Start-Process "\\SoftwareRepo\Citrix\CitrixReceiver.exe" -ArgumentList "/silent /IncludeSSON ADDLOCAL=`"ReceiverInside,ICA_Client,SSON,AM,SELFSERVICE,USB,DesktopViewer,Flash,Vd3d`" ALLOWADDSTORE=A ALLOWSAVEPWD=A ENABLE_SSON=Yes STARTMENUDIR=`"Contoso Citrix`"" -Wait -Verb RunAs
}
else { Write-Host "   Skipped" }

Write-Host "Installing new SSO if necessary ..."
# Install new version of SSO
if (!(Test-Path "C:\Program Files\Sentillion\Agent"))
{
    Start-Process "msiexec" -ArgumentList "/i `"\\SoftwareRepo\Vergence\SSOandCM-x86.msi`" /q ADDLOCAL=DesktopComponents,Authenticator,BridgeWorks,ConfigService VAULTADDRESS=ssovault.contoso.org VAULTSECURITYTOKEN=594681d3-3fe4-4785-bcb5-7b342c3a5899 ALLOWBHOSETTING=YES REBOOT=ReallySuppress" -Wait -Verb RunAs
    #Config Setting for Citrix
    New-ItemProperty -Path "HKLM:\Software\Citrix\AuthManager" -Name ConnectionSecurityMode -Value "Any" -PropertyType string
}
else { Write-Host "   Skipped" }

Write-Host "Installing Rightfax if necessary ..."
#Install RightFax
if (!(Test-Path "C:\Program Files\RightFax") -and $Rightfax -eq $true)
{
	Start-Process "msiexec.exe" -ArgumentList "/i `"\\SoftwareRepo\RightFax\RightFax Product Suite - Client.msi`" /qn REBOOT=ReallySuppress RUNBYRIGHTFAXSETUP=2 CONFIGUREFAXCTRL=1 ADDLOCAL=`"FaxUtil,FaxCtrl,EFM`" INSTALLDIR=`"C:\Program Files\RightFax`" RFSERVERNAME=`"faxserver.contoso.org`"" -Wait -Verb RunAs
    Copy-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Open Text\RightFax FaxUtil.lnk" "$env:PUBLIC\Desktop\RightFax FaxUtil.lnk"
}
else { Write-Host "   Skipped" }

Write-Host "Computer restarting in 30 seconds ..."
Start-Sleep -s 30
Restart-Computer -Force