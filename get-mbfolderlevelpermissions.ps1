Import-Module MSOnline
$O365Cred = Get-Credential 
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Connect-MsolService –Credential $O365Cred


$mailbox = read-host "Enter the Mailbox UPN"


##  Get the permissions at mailbox level
get-MailboxPermission -Identity $Mailbox | ?{$_.user -notlike "NAMPRD*" -and $_.user -notlike "NT AUTHORITY\*" -and $_.user -ne "ExchFullAccessAdmin" -and $_.user -ne "PRDTSB01\JitUsers" -and $_.user -notlike "*View-Only Organization Management"} | select user, accessrights | out-gridview

##  Get the read permission at the root folder (top of information store) and add it to a PSObject $permsObject
$FolderP = get-MailboxFolderPermission -Identity $Mailbox | ?{$_.user.adrecipient -ne $null}

$mailpermissionscollection=@()

$permsObject = New-object PSObject
Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "Folder" -Value ""
Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "User" -Value ""
Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "Rights" -Value ""
        
$permsObject.folder = $FolderP.foldername
$permsObject.user = $FolderP.user
$permsObject.rights = $FolderP.accessrights

$mailpermissionscollection += $permsobject

## Iterate through all folders in mailbox getting permissions that include an entry for an ADRecipient and add that to $permsObject
foreach ($Mfolder in (Get-MailboxFolderStatistics $Mailbox | where { ($_.foldertype -ne "ConversationActions") `
-and ($_.foldertype -notlike "Recoverable*") -and ($_.foldertype -ne "Root") -and ($_.FolderPath -notlike "/Sync*")}))
    {
        $FolderP = $null
        $fname = "$($Mailbox):" + $Mfolder.FolderPath.Replace(“/”,”\”)
        $FolderP = get-MailboxFolderPermission $fname | ?{$_.user.adrecipient -ne $null}

        $permsObject = New-object PSObject
        Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "Folder" -Value ""
        Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "User" -Value ""
        Add-Member -InputObject $permsObject -MemberType NoteProperty -Name "Rights" -Value ""
        
        $permsObject.folder = $FolderP.foldername
        $permsObject.user = $FolderP.user
        $permsObject.rights = $FolderP.accessrights

        $mailpermissionscollection += $permsobject
    }

##  Display all permissions in GridView
$mailpermissionscollection | ?{$_.folder -ne $null} |  out-gridview 

