Get-ADComputer -Filter { OperatingSystem -like '*Windows Server*' } | ForEach-Object {
        $computerName = $_.name
        [PSCustomObject]@{
            ComputerName = $computerName
            ServiceName = ""
            StartingAccount = ""
            StartingMode = ""
        } | Export-Csv .\results.csv -Append -NoTypeInformation

        Get-WmiObject -Class Win32_Service -ComputerName $computerName | select-object name,startname,startmode | ForEach-Object {
            [PSCustomObject]@{
                ComputerName = ""
                ServiceName = $_.name
                StartingAccount = $_.startname
                StartingMode = $_.startmode
            } | Export-Csv .\results.csv -Append -NoTypeInformation
        }
    }