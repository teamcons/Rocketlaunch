


        #===============================================
        #                Initialization                =
        #===============================================


#========================================
# Grab script location in a way that is compatible with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }


#========================================
# Get all resources

# Allow having a fancing GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 


# Load assets
$script:icon                = New-Object system.drawing.icon $ScriptPath\assets\icon.ico
$script:templatefile        = -join($ScriptPath,"\documentation\Project templates.csv")
$script:image               = [system.drawing.image]::FromFile((get-item $ScriptPath\assets\icon-mini.ico))

# Load everything we need
Import-Module $ScriptPath/sources/text.ps1
Import-Module $ScriptPath/sources/defaults.ps1
Import-Module $ScriptPath/sources/internals.ps1
Import-Module $ScriptPath/sources/ui-MainWindow.ps1 
Import-Module $ScriptPath/sources/ui-SettingsDialog.ps1 
Import-Module $ScriptPath/sources/ui-themes.ps1 
Import-Module $ScriptPath/sources/ui-reactiveparts.ps1 
Import-Module $ScriptPath/sources/outlook-backend.ps1 
Import-Module $ScriptPath/sources/main-projectcreation.ps1

# This is the toolbar icon and description
#$GUI_Form_MainWindow.TaskbarItemInfo.Overlay        = $icon
#$GUI_Form_MainWindow.TaskbarItemInfo.Description    = $GUI_Form_MainWindow.Title



        #=======================================================
        #                Display User Interface                =
        #=======================================================



#========================================
# Interface defined in the ui module
#Write-Output "[START] Show main window"; $result = $GUI_Form_MainWindow.ShowDialog()

# Running this without $appContext and ::Run would actually cause a really poor response.
$GUI_Form_MainWindow.Show()


# This makes it pop up
$GUI_Form_MainWindow.Activate()
 
# Create an application context for it to all run within. 
# This helps with responsiveness and threading.
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)


