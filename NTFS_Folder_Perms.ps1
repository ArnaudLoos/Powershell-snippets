######################################################################
# Supply the root folder and the program will recurse through all the sub-folders and write the NTFS permissions to a csv
#
#
######################################################################

# VARIABLES
$ReportPath = "C:\"
$RootFolder = "c:\testperms"        # This is the folder we're getting the permissions of



# Begin Program
$FolderPath = dir -Directory -Path $RootFolder -Recurse -Force
$Report = @()

$Acl = Get-Acl -Path $RootFolder
    foreach ($Access in $acl.Access)
        {
            $Properties = [ordered]@{'FolderName'=$RootFolder;'AD Group or User'=$Access.IdentityReference;'Permissions'=$Access.FileSystemRights;'Inherited'=$Access.IsInherited}
            $Report += New-Object -TypeName PSObject -Property $Properties
        }

Foreach ($Folder in $FolderPath) {
    $Acl = Get-Acl -Path $Folder.FullName
    foreach ($Access in $acl.Access)
        {
            $Properties = [ordered]@{'FolderName'=$Folder.FullName;'AD Group or User'=$Access.IdentityReference;'Permissions'=$Access.FileSystemRights;'Inherited'=$Access.IsInherited}
            $Report += New-Object -TypeName PSObject -Property $Properties
        }
}

$Report | Export-Csv -path $ReportPath\FolderPermissions.csv