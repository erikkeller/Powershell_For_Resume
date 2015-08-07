# CreateDFSFolder

This script was created to automate the process of making a folder in a DFS namespace that was intended to be shared with the entire organization, but only to certain people.

The environment called for these folders to be owned by an employee, who would then approve all other access to this folder. This was meant to "bridge the gap" between users in different departments who may share information as part of a separate committee or group.

Since this was a Windows 7 environment this script ended up on a domain controller (that was also one of the DFS namespace hosts) running Server 2012 r2 since I needed access to the DFS commandlets. This had the side effect of making that DC my preferred remote host for using commandlets that only existed in Windows 8/Server 2012, so WinRM was enabled on that DC along CredSSP, where I would then do a "Enter-PSSession -ComputerName DC-foo -Authentication CredSSP" on my computer when I needed it.

This script was only intended for use on that particular environment and has been scrubbed for identifying information so it won't work as-is without some adjustment.