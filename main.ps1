
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
Add-Type -AssemblyName PresentationCore,PresentationFramework
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 


# Load assets
$script:icon                = New-Object system.drawing.icon $ScriptPath\assets\icon.ico
$script:templatefile        = -join($ScriptPath,"\documentation\Project templates.csv")
$script:image               = [system.drawing.image]::FromFile((get-item $ScriptPath\assets\icon-mini.ico))

# Load everything we need
Import-Module $ScriptPath\sources\text.ps1
Import-Module $ScriptPath\sources\defaults.ps1
Import-Module $ScriptPath\sources\ui-splash.ps1

# We have enough to display the splash
$UI_Splash.Show()
$UI_Splash.Activate()

Import-Module $ScriptPath\sources\internals.ps1
Import-Module $ScriptPath\sources\ui-MainWindow.ps1 
Import-Module $ScriptPath\sources\ui-SettingsDialog.ps1 
Import-Module $ScriptPath\sources\ui-themes.ps1 
Import-Module $ScriptPath\sources\ui-reactiveparts.ps1 

# Outlook Backend is the main intended one for splash
# Because its so fcking slo
Import-Module $ScriptPath\sources\outlook-backend.ps1 
Import-Module $ScriptPath\sources\main-projectcreation.ps1


#========================================
# Interface defined in the ui module
#Write-Output "[START] Show main window"; $result = $GUI_Form_MainWindow.ShowDialog()

# Hide the splash when main UI is shown
# You could just hide it the step before, but i like the idea of a smooth transition
$GUI_Form_MainWindow.Add_Shown({$UI_Splash.Hide()})

# Show the main interface. Thats where everything happens.
# Not ShowDialog because ShowDialog blocks the script and we need appcontext to run
$GUI_Form_MainWindow.Show()


# This makes it pop up
$GUI_Form_MainWindow.Activate()
 
# Create an application context for it to all run within. 
# This helps with responsiveness and threading.
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)


