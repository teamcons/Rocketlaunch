
        #===================================================
        #                GUI - About Dialog                =
        #===================================================


<#
Define the settings interface
Defined but not shown, stuff that on second thought is just clutter

#>




#========================================
Write-Output "[START] Loading graphical user interface"

[int]$GUI_Form_MainWindow_leftalign = 5
[int]$buttonalign = 332



$script:GUI_Form_MoreStuff              = New-Object System.Windows.Forms.Form
$GUI_Form_MoreStuff.Text                = -join($APPNAME," - ",$text.About.tagline)
$GUI_Form_MoreStuff.Icon                = $icon
$GUI_Form_MoreStuff.StartPosition       = "CenterScreen"
$GUI_Form_MoreStuff.Topmost             = $gui_keepontop.Checked
$GUI_Form_MoreStuff.Size                = "340,420"
$GUI_Form_MoreStuff.FormBorderStyle     = "FixedSingle"
$GUI_Form_MoreStuff.MaximizeBox         = $false

# Allow input to window for TextBoxes, etc
#[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($GUI_Form_MoreStuff)

$GUI_Form_MainWindowTabControl                         = New-object System.Windows.Forms.TabControl 
$GUI_Form_MainWindowTabControl.Dock = "Fill" 
$GUI_Form_MoreStuff.Controls.Add($GUI_Form_MainWindowTabControl)



###############################################################################################


Import-Module $MainDir\sources\ui\settings-tab-project.ps1 


###############################################################################################


Import-Module $MainDir\sources\ui\settings-tab-software.ps1 


###############################################################################################

Import-Module $MainDir\sources\ui\settings-tab-about.ps1 
