<#
.SYNOPSIS
    Disables new user sessions, sends message to connected sessions, restarts machine by delivery group
.DESCRIPTION
    Script will place certain machines determined by the grouping parameter in to maintenance mode,
    delay by certain amount of time, warn connected users of an impending restart, and then restart when either
    all users have logged off or a certain amount of time has passed. 
    Note:
    Should work on any machine so long as it has the Citrix PS snapin available, but is safest on a delivery controller.
    Make sure it runs under an account with power and maintenance rights to delivery groups.
    If delay is set, make sure the task is not set to terminate before the delay is over if running as a scheduled task.
.PARAMETER deliveryGroup
    Name of Delivery Group to schedule reboots for. 
    Mandatory.
.PARAMETER timeToWarn
    Time between notification of connected sessions and forced restart in minutes.
    Set to 0 for no warning with immediate restart.
    Default is 30 minutes.
.PARAMETER warning
    The notification to display to connected sessions - can use variable $timeToWarn in message.
    Not displayed if timeToWarn is set to 0.
    Default is "Server restarting in $timeToWarn minutes, please close all Citrix sessions immediately!"
.PARAMETER delay
    Delay in minutes - use to lengthen time between placing machine in maintenance mode and the first warning.
    Default is 0 minutes.
.PARAMETER grouping
    Which machines in the delivery group to restart. Valid inputs are:
    All - All servers in delivery group
    Even - All servers ending in a even number - SERVER NAME MUST END IN A NUMBER!
    Odd - All servers ending in an odd number - SERVER NAME MUST END IN A NUMBER!
    FirstHalf - Divide delivery group in half, all servers in the first half
    SecondHalf - Divide delivery group in half, all servers in the second half
    Default is All.
.PARAMETER logPath
    Defines the location for log files.
    Default is the My Documents folder under the user account running the script.
.PARAMETER adminAddress
    Defines the delivery controller to run the citrix PS commands against.
    Default is localhost.
    
.EXAMPLE
    XenReboot.ps1 -deliveryGroup "Finance"
    Places all machines in the "Finance" delivery group in maintenance mode, restarts any machines
    with no active connections, warns connected users that machine will restart in 30 minutes, then
    restarts after all users have disconnected or 30 minutes have passed.
.EXAMPLE
    XenReboot.ps1 -deliveryGroup "HR" -timeToWarn 0 -delay 360 -grouping Even
    Places even numbered machines in the HR delivery group in to maintenance mode, restarts any
    machines with no active connections, waits 360 minutes (6 hours), then restarts remaining
    machines in maintenance mode immediately with no warning.
.EXAMPLE
    XenReboot.ps1 -deliveryGroup "Maintenance" -timeToWarn 15 -delay 720 -grouping FirstHalf -message "Please sign off within the next $timeToWarn minutes for maintenance restarts."
    Places the first half of the machines in the "Maintenance" delivery group in to maintenance mode,
    restarts any with no connections, waits 720 minutes (12 hours), checks machines for active connections
    and warns those with connections with a custom message, then restarts machines with either 0 active
    connections or until 15 minutes have passed.
.INPUTS
    None. Does not accept piped objects.
.OUTPUTS
    Piped text updates, a csv file showing final status in directory specified by logPath.
.NOTES
    AUTHOR: Erik Keller - ekeller@hammer.net
    TODO: multithreading for large environments, TESTING
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$deliveryGroup,

    [int]$timeToWarn = 30,

    [string]$warning = "Machine restarting in $timeToWarn minutes, please close all Citrix sessions immediately!",

    [int]$delay = 0,
    
    [ValidateSet('All','Even','Odd','FirstHalf','SecondHalf')]
    [string]$grouping='All',

    [string]$logPath = [environment]::getfolderpath("mydocuments"),

    [string]$adminAddress = "localhost"
)

# Static Variables
$REGWINDOW = 30
$LOGOFFWAIT = 10

Add-PSSnapin Citrix.* -ErrorAction SilentlyContinue
if ($adminAddress -ne "localhost")
{
    $servers = Get-BrokerMachine -DesktopGroupName $deliveryGroup -InMaintenanceMode $False -PowerState On -MaxRecordCount 5000 -AdminAddress $adminAddress
}
else 
{
    $servers = Get-BrokerMachine -DesktopGroupName $deliveryGroup -InMaintenanceMode $False -PowerState On -MaxRecordCount 5000
}
$targetServers = @()

Write-Output "XenReboot script started at $(Get-Date -Format F)"
Write-Output "Options chosen:"
Write-Output "  Delivery Group - $deliveryGroup"
if ($timeToWarn -gt 0)
{
    Write-Output "  Time between warning and forced restart - $timeToWarn Minutes."
    Write-Output "  Warning sent - $warning"
}
else 
{
    Write-Output "  No warning before restart."
}
if ($delay -gt 0)
{
    Write-Output "  Delay between maintenance mode and warning/restart - $delay Minutes."
}
Write-Output "  Group to restart - $grouping."
Write-Output "  Path for logging - $logPath"

# Build the targetServers variable with machines we want to reboot
Switch($grouping)
{
    'All'
    {
        Foreach($server in $servers)
        {
            $targetServers += New-Object PSObject -Property @{
                    Server = $server
                    online = $true
                    inMaintenanceMode = $False
                    rebooted = $False
                    timeWarned = $null
                    timeRebooted = $null
                    timeDeadline = $null
                    nextAction = "PREP"
                    sessionCount = 1
            }
        } 
    }
    'Even'
    {
        Foreach($server in $servers)
        {
            $serverName = $server.HostedMachineName
            $number = ([regex]::Match($serverName,'.*(\d{1})\w{0,2}$')).captures.Groups[1].value -as [int]

            If (($number % 2) -eq 0) 
            { 
                $targetServers += New-Object PSObject -Property @{
                    Server = $server
                    online = $true
                    inMaintenanceMode = $False
                    rebooted = $False
                    timeWarned = $null
                    timeRebooted = $null
                    timeDeadline = $null
                    nextAction = "PREP"
                    sessionCount = 1
                }
            }
        }
    }
    'Odd'
    {
        Foreach($server in $servers)
        {
            $serverName = $server.HostedMachineName
            $number = ([regex]::Match($serverName,'.*(\d{1})\w{0,2}$')).captures.Groups[1].value -as [int]

            If (($number % 2) -ne 0) 
            {
                $targetServers += New-Object PSObject -Property @{
                    Server = $server
                    online = $true
                    inMaintenanceMode = $False
                    rebooted = $False
                    timeWarned = $null
                    timeRebooted = $null
                    timeDeadline = $null
                    nextAction = "PREP"
                    sessionCount = 1
                } 
            }
        }
    }
    'FirstHalf'
    {
        For ($x = 0; $x -lt ($servers.Count / 2); $x++)
        {
            $targetServers += New-Object PSObject -Property @{
                    Server = $servers[$x]
                    online = $true
                    inMaintenanceMode = $False
                    rebooted = $False
                    timeWarned = $null
                    timeRebooted = $null
                    timeDeadline = $null
                    nextAction = "PREP"
                    sessionCount = 1
            }
        }
    }
    'SecondHalf'
    {
        For ($x = ($servers.Count / 2); $x -lt ($servers.Count); $x++)
        {
            $targetServers += New-Object PSObject -Property @{
                    Server = $servers[$x]
                    online = $true
                    inMaintenanceMode = $False
                    rebooted = $False
                    timeWarned = $null
                    timeRebooted = $null
                    timeDeadline = $null
                    nextAction = "PREP"
                    sessionCount = 1
            }
        }
    }
}

# Reboots the machine
function Reboot 
{
    New-BrokerHostingPowerAction -MachineName $target.Server.DNSName -Action Restart | Write-Debug
    $target.rebooted = $True
    $target.timeRebooted = Get-Date
    $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes $REGWINDOW)
    $target.online = $False
    Write-Output "$($target.Server.MachineName) restarted at $(Get-Date -format T)"
    #Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $False
    #$target.inMaintenanceMode = $False
    #Write-Output "$($target.Server.MachineName) removed from maintenance mode at $(Get-Date -format T)"
}

# Preps machine for reboot - used when $delay is greater than 0
function PrepServerDelay
{
    $target.sessionCount = (Get-BrokerMachine -DesktopGroupName $deliveryGroup -DNSName $target.Server.DNSName).SessionCount
    Write-Output "$($target.Server.MachineName) checked for connected users at $(Get-Date -format T), found $($target.sessionCount)"

    if ($target.sessionCount -eq 0)
    {
        Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $True
        $target.inMaintenanceMode = $True
        $target.nextAction = "POSTREBOOT"
        Write-Output "$($target.Server.MachineName) placed in to maintenance mode at $(Get-Date -format T)"
        Reboot
    }
    else 
    {
        Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $True
        $target.inMaintenanceMode = $True
        Write-Output "$($target.Server.MachineName) placed in to maintenance mode at $(Get-Date -format T)"

        if ($timeToWarn -eq 0)
        {
            $target.timeDeadline = Get-Date
            $target.nextAction = "LOGOFF"
        }
        else 
        {
            $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes ($timeToWarn + $delay))
            $target.nextAction = "WARN"
        }
    }
}

# Preps machine for reboot - used when $delay is 0
function PrepServerNoDelay
{
    $target.sessionCount = (Get-BrokerMachine -DesktopGroupName $deliveryGroup -DNSName $target.Server.DNSName).SessionCount
    Write-Output "$($target.Server.MachineName) checked for connected users at $(Get-Date -format T), found $($target.sessionCount)"

    if ($target.sessionCount -eq 0)
    {
        Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $True
        $target.inMaintenanceMode = $True
        $target.nextAction = "POSTREBOOT"
        Write-Output "$($target.Server.MachineName) placed in to maintenance mode at $(Get-Date -format T)"
        Reboot
    }
    else 
    {
        Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $True
        $target.inMaintenanceMode = $True
        Write-Output "$($target.Server.MachineName) placed in to maintenance mode at $(Get-Date -format T)"
        
        if ($timeToWarn -gt 0)
        {
            Get-BrokerSession -HostedMachineName $target.Server.HostedMachineName | Send-BrokerSessionMessage -Title "Attention!" -Text $warning -MessageStyle Exclamation
            $target.timeWarned = Get-Date
            $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes $timeToWarn)
            $target.nextAction = "LOGOFF"
            Write-Output "$($target.Server.MachineName) warned connected users at $(Get-Date -format T)"
        }
        else 
        {
            $target.timeDeadline = Get-Date
            $target.nextAction = "LOGOFF"
        }
    }
}

# Checks if server is ready to reboot and reboots if it is
function CheckServer
{
    if ($target.nextAction -eq "WARN")
    {
        Get-BrokerSession -HostedMachineName $target.Server.HostedMachineName | Send-BrokerSessionMessage -Title "Attention!" -Text $warning -MessageStyle Exclamation
        $target.timeWarned = Get-Date
        $target.nextAction = "LOGOFF"
        $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes 1)
        Write-Output "$($target.Server.MachineName) warned connected users at $(Get-Date -format T)"
    }
    elseif ($target.nextAction -eq "LOGOFF" -and $target.timeDeadline -lt (Get-Date))
    {
        Get-BrokerSession -MachineName $target.Server.Machinename | Stop-BrokerSession
        Write-Output "$($target.Server.MachineName) logged off all users at $(Get-Date -format T)"
        $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes $LOGOFFWAIT)
        $target.nextAction = "REBOOT"
    }
    elseif ($target.nextAction -eq "REBOOT" -and $target.timeDeadline -lt (Get-Date))
    {
        Reboot
        $target.nextAction = "POSTREBOOT"
        $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes $REGWINDOW)
    }
    elseif ($target.nextAction -eq "POSTREBOOT" -and $target.timeDeadline -gt (Get-Date))
    {
        Write-Output "$($target.Server.MachineName) checked for post reboot status at $(Get-Date -format T)"
        $registrationState = (Get-BrokerMachine -DNSName $target.Server.DNSName).RegistrationState

        if ($registrationState -eq "Registered")
        {
            Write-Output "$($target.Server.MachineName) found online and registered."
            $target.online = $True
            Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $False
            $target.inMaintenanceMode = $False
            $target.nextAction = "FINISHED"
        }
        else 
        {
            Write-Output "$($target.Server.MachineName) status is $registrationState"
        }
    }
    elseif ($target.nextAction -eq "POSTREBOOT" -and $target.timeDeadline -lt (Get-Date))
    {
        Write-Output "$($target.Server.MachineName) checked for post reboot status at $(Get-Date -format T)"
        $registrationState = (Get-BrokerMachine -DNSName $target.Server.DNSName).RegistrationState

        if ($registrationState -eq "Registered")
        {
            Write-Output "$($target.Server.MachineName) found online and registered."
            $target.online = $True
            Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $False
            $target.inMaintenanceMode = $False
            $target.nextAction = "FINISHED"
        }
        else
        {
            Write-Output "$($target.Server.MachineName) found offline and has passed its registration window."
            New-BrokerHostingPowerAction -MachineName $target.Server.DNSName -Action TurnOn | Write-Debug
            Write-Output "$($target.Server.MachineName) manually restarted."
            $target.timeDeadline = (Get-Date) + (New-TimeSpan -Minutes $REGWINDOW)
            $target.nextAction = "MANUALSTART"
        }
    }
    elseif ($target.nextAction -eq "MANUALSTART" -and $target.timeDeadline -gt (Get-Date))
    {
        Write-Output "$($target.Server.MachineName) checked for post manual start status at $(Get-Date -format T)"
        $registrationState = (Get-BrokerMachine -DNSName $target.Server.DNSName).RegistrationState

        if ($registrationState -eq "Registered")
        {
            Write-Output "$($target.Server.MachineName) found online and registered."
            $target.online = $True
            Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $False
            $target.inMaintenanceMode = $False
            $target.nextAction = "FINISHED"
        }
        else 
        {
            Write-Output "$($target.Server.MachineName) status is $registrationState"
        }
    }
    elseif ($target.nextAction -eq "MANUALSTART" -and $target.timeDeadline -lt (Get-Date))
    {
        Write-Output "$($target.Server.MachineName) checked for post manual start status at $(Get-Date -format T)"
        $registrationState = (Get-BrokerMachine -DNSName $target.Server.DNSName).RegistrationState

        if ($registrationState -eq "Registered")
        {
            Write-Output "$($target.Server.MachineName) found online and registered."
            $target.online = $True
            Set-BrokerMachine $target.Server.MachineName -InMaintenanceMode $False
            $target.inMaintenanceMode = $False
            $target.nextAction = "FINISHED"
        }
        else 
        {
            Write-Output "$($target.Server.MachineName) found offline and has passed its registration window. Something has gone wrong."
            $target.nextAction = "FINISHED"
        }
    }
}

# Prep machines
foreach ($target in $targetServers)
{
    if ($delay -gt 0)
    {
        PrepServerDelay
    }
    else 
    {
        PrepServerNoDelay
    }
}

# Sleep for delay if greater than 0
Write-Output "XenReboot script sleeping for $($delay + 1) minutes."
Start-Sleep -Seconds ($delay * 60 + 60)

# Check the servers until all have restarted
while (($targetServers | Where-Object { $_.rebooted -eq $False -or $_.online -eq $false } | Measure-Object).Count -gt 0)
{
    foreach ($target in $targetServers)
    {
        if ($target.rebooted -eq $False -or $target.online -eq $false)
        {
            CheckServer
            Start-Sleep -Seconds 30
        } 
    }
}

Write-Output "XenReboot script finished at $(Get-Date -format F)"
$targetServers | Export-CSV -NoTypeInformation ($logPath + "XenRebootReport.csv")