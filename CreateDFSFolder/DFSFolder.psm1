<# 
 .Synopsis
  Creates a shared folder in a set DFS namespace for a certain organization 

 .Description
  Takes input of folder name and file server host for the folder and creates a DFS
  folder in the X namespace, creates associated Read and Full access groups,
  and adds the groups to explicit view permission.

 .Parameter FolderName
  Name of DFS Folder

 .Parameter FileSharePath
  Share to link the DFS folder to - will be created if it doesn't already exist

 .Example
   # Create a folder on the # drive called "Clinical Data" on file server FILESERVER23
   New-DFSFolder -FolderName "Clinical Data" -FileSharePath "\\FILESERVER23\Folder\ClinData"
#>

function New-DFSFolder
{
	param
	(
		[string]$FolderName,
		[string]$FileSharePath
	)

	Try
	{
		New-ADGroup -Name "Data_$($FolderName)_R" -SamAccountName "Data_$($FolderName)_R" -GroupCategory Security -GroupScope Global -DisplayName "Data $($FolderName) Read" -Path "OU=Access Group,OU=Corporate Groups,DC=CONTOSO,DC=ORG" -Description "Read only access to folder $FolderName" -PassThru | Add-ADPrincipalGroupMembership -MemberOf "Data_DFS_#"
		New-ADGroup -Name "Data_$($FolderName)_Full" -SamAccountName "Data_$($FolderName)_Full" -GroupCategory Security -GroupScope Global -DisplayName "Data $($FolderName) Full" -Path "OU=Access Group,OU=Corporate Groups,DC=CONTOSO,DC=ORG" -Description "Full access to folder $FolderName" -PassThru | Add-ADPrincipalGroupMembership -MemberOf "Data_DFS_#"
		
		if (!(Test-Path "$FileSharePath")) { New-Item -Path "$FileSharePath" -ItemType Directory }
		
		#ACL Block
		$readRights = [System.Security.AccessControl.FileSystemRights]"Read, ExecuteFile"
		$allRights = [System.Security.AccessControl.FileSystemRights]"Read, ExecuteFile, Modify, Write"
		#Inheritance flag is a bit flag, need to combine the flags for ObjectInherit and ContainerInherit with a bitwise OR to do both
		$inheritFlag = [System.Security.AccessControl.InheritanceFlags]::ObjectInherit -bor [System.Security.AccessControl.InheritanceFlags]::ContainerInherit
		#Already set the inheritance flags so propogation isn't necessary
		$propagateFlag = [System.Security.AccessControl.PropagationFlags]::None
		#Check allow for rights specified
		$objType = [System.Security.AccessControl.AccessControlType]::Allow
		#Users to set rights to
		$readUser = New-Object System.Security.Principal.NTAccount("CONTOSO\Data_$($FolderName)_R")
		$allUser = New-Object System.Security.Principal.NTAccount("CONTOSO\Data_$($FolderName)_Full")
		#Create and add ACLs to folder object
		$readACL = New-Object System.Security.AccessControl.FileSystemAccessRule ($readUser, $readRights, $inheritFlag, $propagateFlag, $objType)
		$allACL = New-Object System.Security.AccessControl.FileSystemAccessRule ($allUser, $allRights, $inheritFlag, $propagateFlag, $objType)
		$ACL = Get-Acl "$FileSharePath"
		$ACL.AddAccessRule($readACL)
		$ACL.AddAccessRule($allACL)
		Set-Acl -Path "$FileSharePath" -AclObject $ACL
		#End ACL Block

		New-DfsnFolder -Path "\\CONTOSO\DFSPath\$FolderName" -TargetPath "$FileSharePath"
		dfsutil property sd grant \\CONTOSO\DFSPath\$FolderName CONTOSO\data_$($FolderName)_R:RX protect
		dfsutil property sd grant \\CONTOSO\DFSPath\$FolderName CONTOSO\data_$($FolderName)_Full:RX protect
	}
	catch [UnauthorizedAccessException]
	{
		Write-Host "Missing access rights : $($_.Exception.Message)"
	}
	catch [IOException]
	{
		Write-Host "Unable to create folder $FileServer : $($_.Exception.Message)"
	}
	catch
	{
		Write-Host "Unhandled exception error : $($_.Exception.Message)"
	}
}
Export-ModuleMember -Function New-DFSFolder