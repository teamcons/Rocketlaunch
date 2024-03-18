

# Shenanigans to function well with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }



# Folder on the desktop. Not touching Program Files im not crazi
$WHERE = [Environment]::GetFolderPath("Desktop")

# How folder is called
$DIRNAME = $(Get-Item $ScriptPath).Name



# Copy the whole folder there
Copy-Item -Path $ScriptPath -Destination $WHERE


# Define how to create shortcuts
function Create-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutPath
    )
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}



# Path to executable
$EXE = -join($WHERE,"\",$DIRNAME,"\Start.Rocketlaunch.exe")

# Lets fight Ganon
$LINK = -join($WHERE,"\Start.Rocketlaunch.lnk")

# Create link to exe
Create-Shortcut -TargetPath $EXE -ShortcutPath $LINK




# Pin shortcuts to the taskbar
$shell = New-Object -ComObject Shell.Application
$taskbarPath = [System.IO.Path]::Combine([Environment]::GetFolderPath('ApplicationData'), 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar')
$shell.Namespace($taskbarPath).Self.InvokeVerb('pindirectory',$LINK)


