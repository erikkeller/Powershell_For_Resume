<# PrinterCheckAndClear.ps1

Does a remote PSSession and registry dive and checks for blank or missing keys in
the list of installed printer drivers, then deletes the root key to
force a re-download of the print driver. Restarts the print spooler and 
runs gpupdate /force on the remote PC after it is done.

Version 1.00 - Initial Creation - EK 09162014
Version 1.01 - Commented with a vengeance - EK 09172014
Version 1.02 - Fixed bug where the regkey count would come 
               back as 1 instead of 0 for some bad keys. Deletes
               the Sharpdesk Composer key too but literally nobody 
               uses that anyway - EK 10152014
#>

###################################### 
# Check for elevated rights
#
# Checks to make sure the script is running with the admin security principal
# If not, starts a new shell as admin running the same script and ends the current shell
#
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}
#
# End check for elevated rights
######################################

######################################
# Script block to run on remote computer
#
# Script block that gets passed to the remote session started on the affected PC
#
$scriptBlock = {

    # Needs path to driver registry keys
    Param ([string]$rootpath)

    # Get list of drivers stored in printer driver registry key
    # Syntax is important due to PS 2.0, -Name flag must exist
    $drivers = Get-ChildItem -Path $rootpath -Name
    #Write-Host "Drivers: $drivers" # Debug code
    foreach ($driver in $drivers)
    {
        #Write-Host "Driver: $driver" # Debug code
        $fixedPath = $rootpath + "\" + $driver
        #Write-Host "FixedPath: $fixedPath" # Debug code

        # Script would hang on a bad path occasionally
        if (Test-Path $fixedPath) 
        { 
            $regKey = (Get-ItemProperty -Path $fixedPath).'Dependent Files'
            #Write-Host "Key count: $($regKey.Count)" # Debug code

            # If the Dependent Files key is blank, delete the driver key
            if ($regKey.Count -le 1)
            {
                Remove-Item $fixedPath -Recurse
                Write-Host "Removing $driver registry key"
            }
        }
    }
    # Duh
    Write-Host "Restarting Print Spooler..."
    Restart-Service spooler -Force
    Write-Host "Forcing Group Policy Update..."
    gpupdate /force /wait:0
    # End duh
}
#
# End script block to run on remote computer
##################################

##################################
# Globals
#
$hostname = Read-Host "Hostname to check"
$rootpath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers\Version-3"
#
# End Globals
##################################

# Check for computer connectivity
if (Test-Connection -ComputerName $hostname -Count 2 -Quiet)
{
    # Start a new remote session
    $session = New-PSSession -ComputerName $hostname
    # Run the script block on the remote session
    $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $rootpath -Verbose
    # End the remote session
    Remove-PSSession $session
    # Display any output that didn't output interactively
    Write-Host "`n$result"
}
else # Host didn't respond
{
    Write-Host "Host $hostname did not respond, check the hostname and try again"
}

# Powershell doesn't have a "pause" command so we reinvent our own wheel
Write-Host "Press any key to continue..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")