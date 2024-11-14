
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

$global:MainDir = $ScriptPath

#========================================
# Get all resources

# Allow having a fancing GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore,PresentationFramework

# Ensure companion scripts run
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy ByPass -F

# Load and parse the JSON configuration file
$script:settings = (Get-Content $MainDir\data\settings.json -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue)


# Get localization
# If it is default, revert to system language
if ($settings.UI.Language -match "System")
{
    $script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization
}
# Else take whatever is indicated
else {
    $script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization -UICulture $settings.UI.Language
}



[System.Windows.Forms.Application]::EnableVisualStyles() 


Import-Module $MainDir\sources\defaults.ps1
Import-Module $MainDir\sources\ui\splash.ps1
Import-Module $MainDir\sources\internals.ps1
Import-Module $MainDir\sources\ui\mainWindow.ps1 
Import-Module $MainDir\sources\ui\settingsDialog.ps1 
Import-Module $MainDir\sources\ui\themes.ps1 
Import-Module $MainDir\sources\ui\reactiveparts.ps1 



Change-Theme $settings.UI.Theme

# We have enough to display the splash
$UI_Splash.Show()
$UI_Splash.Activate()

# Outlook Backend is the main intended one for splash
# Because its so fcking slo
Import-Module $MainDir\sources\outlook-backend.ps1 
Import-Module $MainDir\sources\main-projectcreation.ps1






#========================================
# Interface defined in the ui module
#Write-Output "[START] Show main window"; $result = $GUI_Form_MainWindow.ShowDialog()

# Hide the splash when main UI is shown
# You could just hide it the step before, but i like the idea of a smooth transition
$GUI_Form_MainWindow.Add_Shown({$UI_Splash.Hide() })

# Show the main interface. Thats where everything happens.
# Not ShowDialog because ShowDialog blocks the script and we need appcontext to run
$GUI_Form_MainWindow.Show()


# This makes it pop up
$GUI_Form_MainWindow.Activate()
 
# Create an application context for it to all run within. 
# This helps with responsiveness and threading.
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)


