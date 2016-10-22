<# CitrixShortcuts.psm1

    Functions to save and restore Citrix shortcuts placed in the Desktop,
    Taskbar, and Pinned Start Menu

    Defaults are for currently logged in User
#>

function Save-CitrixShortcuts
{
    [CmdletBinding()]
    param(
        [Parameter(Position=1)]     #User's DESKTOP path
        [string]$userDesktop = "$env:HOMEDRIVE\$env:HOMEPATH\Desktop",

        [Parameter(Position=2)]     #User's ROAMING APPDATA path
        [string]$userAppData = $env:APPDATA,

        [Parameter(Position=3)]     #Location to save shortcut reference file
        [string]$saveLocation = "$env:HOMEDRIVE\$env:HOMEPATH"
    )

    $desktopShortcuts = @()
    $taskbarShortcuts = @()
    $startShortcuts = @()
    $shell = New-Object -COM WScript.Shell

    Get-ChildItem $("$userDesktop\*.lnk") | % {
        $shortcut = $shell.CreateShortcut("$($_.FullName)")

        if ($shortcut.TargetPath -like "*SelfService.exe")
        {
            $table = New-Object PSObject -Property @{
                FullName            = $shortcut.FullName
                Arguments           = $shortcut.Arguments
                Description         = $shortcut.Description
                IconLocation        = $shortcut.IconLocation
                TargetPath          = $shortcut.TargetPath
                WindowStyle         = $shortcut.WindowStyle
                WorkingDirectory    = $shortcut.WorkingDirectory
            }
            $desktopShortcuts += , $table
        }
    }

    Get-ChildItem $("$userAppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\Taskbar\*.lnk") | % {
        $shortcut = $shell.CreateShortcut("$($_.FullName)")

        if ($shortcut.TargetPath -like "*SelfService.exe")
        {
            $table = New-Object PSObject -Property @{
                FullName            = $shortcut.FullName
            }
            $taskbarShortcuts += , $table
        }
    }

    Get-ChildItem $("$userAppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\*.lnk") | % {
        $shortcut = $shell.CreateShortcut("$($_.FullName)")

        if ($shortcut.TargetPath -like "*SelfService.exe")
        {
            $table = New-Object PSObject -Property @{
                FullName            = $shortcut.FullName
            }
            $startShortcuts += , $table
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Write-Debug
    $desktopShortcuts | Export-CSV "$saveLocation\SavedDesktopShortcuts.csv" -NoTypeInformation -Force
    $taskbarShortcuts | Export-CSV "$saveLocation\SavedTaskbarShortcuts.csv" -NoTypeInformation -Force
    $startShortcuts | Export-CSV "$saveLocation\SavedStartShortcuts.csv" -NoTypeInformation -Force
}

function Restore-CitrixShortcuts
{
    [CmdletBinding()]
    Param(
        [Parameter(Position=1)]     #Location of saved shortcuts
        [string]$savedShortcutsLocation = "$env:HOMEDRIVE\$env:HOMEPATH",

        [Parameter(Position=2)]     #File location of start menu
        [string]$startMenuLocation = [environment]::GetFolderPath("StartMenu")
    )

    $desktopShortcuts = Import-CSV "$savedShortcutsLocation\SavedDesktopShortcuts.csv"
    $taskbarShortcuts = Import-CSV "$savedShortcutsLocation\SavedTaskbarShortcuts.csv"
    $startShortcuts = Import-CSV "$savedShortcutsLocation\SavedStartShortcuts.csv"
    $shell = New-Object -COM WScript.Shell

    $desktopShortcuts | % {
        $shortcut                   = $shell.CreateShortcut("$($_.FullName)")
        $shortcut.Arguments         = $_.Arguments
        $shortcut.Description       = $_.Description

        if (Test-Path($_.IconLocation)) 
        { 
            $shortcut.IconLocation = $_.IconLocation 
        }
        else 
        {
            $name = ([regex]::Match($(Split-Path $_.IconLocation),'(\w+)(?:_\d+\.ico,0)')).captures.Groups[2].value -as [string]
            $newIcon = Get-ChildItem "$env:APPDATA\Citrix\SelfService\Icons" | where { $_.Name -like "*$name*" }
            $shortcut.IconLocation = $newIcon.FullName + ",0"
        }

        $shortcut.TargetPath        = $_.TargetPath
        $shortcut.WindowStyle       = $_.WindowStyle
        $shortcut.WorkingDirectory  = $_.WorkingDirectory
        $shortcut.Save()
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Write-Debug
    $shell = New-Object -COM Application.Shell

    $taskbarShortcuts | % {
        $name = (Split-Path $_.FullName -Leaf).Replace(".lnk","")
        Get-ChildItem $startMenuLocation -Recurse | Where { $_.Name -like "*$name*" } | % {
            $pin = $shell.Namespace($(Split-Path $_.FullName)).Parsename($(Split-Path $_.FullName -leaf))
            $pin.invokeverb('taskbarpin')
        }
    }

    $startShortcuts | % {
        $name = (Split-Path $_.FullName -Leaf).Replace(".lnk","")
        Get-ChildItem $startMenuLocation -Recurse | Where { $_.Name -like "*$name*" } | % {
            $pin = $shell.Namespace($(Split-Path $_.FullName)).Parsename($(Split-Path $_.FullName -leaf))
            $pin.invokeverb('startpin')
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Write-Debug
}

Export-ModuleMember -Function 'Save-CitrixShortcuts'
Export-ModuleMember -Function 'Restore-CitrixShortcuts'