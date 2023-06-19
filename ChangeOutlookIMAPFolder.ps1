#author:        Manuel Würth
#company:       avatex solutions GmbH
#date:          18.06.2023
#description:   If you migrate a IMAP mailbox to exchange online with a pst, it can happen that the foldertype of the migrated folders is wrong.
#               Start the script an choose the folder you would like to migrate. You can migrate the subfolders too.


$logfile = "{0}\MigrateImapFolder.log" -f $env:TEMP

function Log {
    param(
        [string]$Message,
        [string]$type
    )
    try {
        $nowdate = Get-Date -format "dd.MM.yyyy HH:mm"

        if (!(Test-Path $logfile)) {
            $ret = Set-Content -Path $logfile -Value ($nowdate + "- Info - Log created")
        }

        $LogMessage = "{0} - {1} - {2}" -f $nowdate, $type, $Message
        $ret = Add-Content -Path $logfile -Value $LogMessage
    }
    catch {
        #do nothing
    }
}

function MigrateFolder {
    param(
        [System.__ComObject]$FolderToFix
    )
    try {
        $PropName = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
        $Value = "IPF.Note"
    
        $propAccessor = $FolderToFix.PropertyAccessor
        $propAccessorProperty = $propAccessor.GetProperty($PropName)
        if ($propAccessorProperty -eq "IPF.Imap") {
            $ret = $propAccessor.SetProperty($PropName, $Value)
            $ret = Log -Message ("Foldertype of folder " + $FolderToFix.Name + " changed to IPF.Note") -type "Info"
        }
        else {
            $ret = Log -Message ("Foldertype of folder " + $FolderToFix.Name + " is already IPF.Note. No migration needed.") -type "Info"
        }       
    }
    catch {
        $ret = Log -Message ("Error while trying to migrate folder " + $FolderToFix.Name) -type "Error"
        $ret = Log -Message $_.Exception.Message -type "Error"
    }
}

function Main {

    $ret = Log -Message "************** Script start **************" -type "Info"

    Add-Type -AssemblyName PresentationCore, PresentationFramework

    try {
        $objOutlook = New-Object -ComObject "Outlook.Application"
        $FolderToFix = $objOutlook.Application.Session.PickFolder()
        if ($FolderToFix) {
            $ret = Log -Message ("Folder " + $FolderToFix.Name + " choosed") -type "Info"
        }
        else {
            $ret = Log -Message ("No Folder choosed. Exit script.") -type "error"
        }
    }
    catch {
        $ret = Log -Message "Error while trying to create outlook comobject." -type "Error"
        $ret = Log -Message $_.Exception.Message -type "Error"
    }

    $MigrateSubfolders = [System.Windows.MessageBox]::Show("Convert subfolders?", "Migrate IMAP folders", [System.Windows.MessageBoxButton]::YesNo)

    If ($MigrateSubfolders -eq 6) {

        $ret = Log -Message "Migrate subfolders" -type "Info"
        $ret = MigrateFolder -FolderToFix $FolderToFix

        foreach ($folder in $FolderToFix.Folders) {
            $ret = MigrateFolder -FolderToFix $folder
        }
    }
    else {
        $ret = Log -Message "Don't migrate subfolders" -type "Info"
        $ret = MigrateFolder -FolderToFix $FolderToFix
    }
}

Main

$ShowLogfile = [System.Windows.MessageBox]::Show("Show logfile?", "Migrate IMAP folders", [System.Windows.MessageBoxButton]::YesNo)

if ($ShowLogfile -eq 6) {
    $ret = Start-Process -FilePath "notepad.exe" -ArgumentList $logfile
}
