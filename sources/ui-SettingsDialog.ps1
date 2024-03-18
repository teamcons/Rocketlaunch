
        #===================================================
        #                GUI - About Dialog                =
        #===================================================

Write-Output "[START] Loading graphical user interface"


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

[int]$GUI_Form_MainWindow_leftalign = 5

$script:GUI_Form_MoreStuff              = New-Object System.Windows.Forms.Form
$GUI_Form_MoreStuff.Text                = -join($APPNAME," - ",$text_aboutsubtitle)
$GUI_Form_MoreStuff.Icon                = $icon

$GUI_Form_MoreStuff.StartPosition       = "CenterScreen"
$GUI_Form_MoreStuff.Topmost             = $true
$GUI_Form_MoreStuff.Size                = "340,420"
$GUI_Form_MoreStuff.FormBorderStyle     = "FixedSingle"
$GUI_Form_MoreStuff.MaximizeBox         = $false

$GUI_Form_MainWindowTabControl                         = New-object System.Windows.Forms.TabControl 
$GUI_Form_MainWindowTabControl.Dock = "Fill" 
$GUI_Form_MoreStuff.Controls.Add($GUI_Form_MainWindowTabControl)

####################################

$GUI_Tab_Settings = New-object System.Windows.Forms.Tabpage
$GUI_Tab_Settings.Name = "Advanced" 
$GUI_Tab_Settings.Text = $text_settingstag
$GUI_Tab_Settings.UseVisualStyleBackColor = $True 

# Label above input
$moresettingstitle                     = New-Object System.Windows.Forms.Label
$moresettingstitle.Size                = New-Object System.Drawing.Size(300,20)
$moresettingstitle.Left                = $GUI_Form_MainWindow_leftalign
$moresettingstitle.Top                 = 10
$moresettingstitle.Text                = $text_settingstag
$moresettingstitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$moresettingsnota                     = New-Object System.Windows.Forms.Label
$moresettingsnota.Size                = New-Object System.Drawing.Size(300,20)
$moresettingsnota.Left                = $GUI_Form_MainWindow_leftalign
$moresettingsnota.Top                 = 30
$moresettingsnota.Text                = $text_settingsnota
$moresettingsnota.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 7, [System.Drawing.FontStyle]::Italic)


$CheckIfCreateExplorerQuickAccess                = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateExplorerQuickAccess.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,60)
$CheckIfCreateExplorerQuickAccess.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfCreateExplorerQuickAccess.Text           = $text_settings_ExplorerQuickAccess
$CheckIfCreateExplorerQuickAccess.Checked        = $default_createshortcut

$CheckIfCreateOutlookFolder                = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateOutlookFolder.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,90)
$CheckIfCreateOutlookFolder.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfCreateOutlookFolder.Text           = $text_settings_OutlookFolder
$CheckIfCreateOutlookFolder.Checked        = $default_createoutlookfolder



$CheckIfExpandArchives                = New-Object System.Windows.Forms.CheckBox        
$CheckIfExpandArchives.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,120)
$CheckIfExpandArchives.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfExpandArchives.Text           = $text_settings_ExpandArchives
$CheckIfExpandArchives.Checked        = $default_expandarchives


$CheckIfCountWords                = New-Object System.Windows.Forms.CheckBox        
$CheckIfCountWords.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,150)
$CheckIfCountWords.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfCountWords.Text           = $text_settings_Countwords
$CheckIfCountWords.Checked        = $default_countwords


$CheckIfOpenExplorer                = New-Object System.Windows.Forms.CheckBox        
$CheckIfOpenExplorer.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,180)
$CheckIfOpenExplorer.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfOpenExplorer.Text           = $text_settings_OpenExplorer
$CheckIfOpenExplorer.Checked        = $default_openexplorer

$CheckIfNotify                = New-Object System.Windows.Forms.CheckBox        
$CheckIfNotify.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,210)
$CheckIfNotify.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfNotify.Text           = $text_settings_Notify
$CheckIfNotify.Checked        = $default_notifywhenfinished

$helptitle                     = New-Object System.Windows.Forms.Label
$helptitle.Size                = New-Object System.Drawing.Size(300,20)
$helptitle.Left                = $GUI_Form_MainWindow_leftalign
$helptitle.Top                 = 280
$helptitle.Text                = $text_settings_helptitle
$helptitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$getthedoc                 = New-Object System.Windows.Forms.Button
$getthedoc.Size            = New-Object System.Drawing.Size (180,30)
$getthedoc.Left            = $GUI_Form_MainWindow_leftalign
$getthedoc.Top             = 310
$getthedoc.Text            = $text_settings_getthedoc
#$getthedoc.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch/raw/main/docs/Manual%20-%20Rocketlaunch.docx"})
$getthedoc.Add_Click( {start-process '$ScriptPath\..\documentation\Rocketlaunch Manual.docx' } )

$askme                 = New-Object System.Windows.Forms.Button
$askme.Size            = New-Object System.Drawing.Size (120,30)
$askme.Left            = ($GUI_Form_MainWindow_leftalign + 190)
$askme.Top             = 310
$askme.Text            = $text_settings_askme
$askme.Add_Click( {start-process "Mailto:stella.menier@gmx.de"})

$GUI_More_Close                               = New-Object System.Windows.Forms.Button
$GUI_More_Close.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign ),140)
$GUI_More_Close.Size                          = New-Object System.Drawing.Size(120,25)
$GUI_More_Close.Text                          = $text_close_settings
$GUI_More_Close.UseVisualStyleBackColor       = $True

$GUI_Tab_Settings.Controls.Add($moresettingstitle)
$GUI_Tab_Settings.Controls.Add($moresettingsnota)
#$GUI_Tab_Settings.Controls.Add($GUI_More_Close)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateExplorerQuickAccess)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateOutlookFolder)
$GUI_Tab_Settings.Controls.Add($CheckIfExpandArchives)
$GUI_Tab_Settings.Controls.Add($CheckIfCountWords)
$GUI_Tab_Settings.Controls.Add($CheckIfOpenExplorer)
$GUI_Tab_Settings.Controls.Add($CheckIfNotify)
$GUI_Tab_Settings.Controls.Add($helptitle)
$GUI_Tab_Settings.Controls.Add($getthedoc)
$GUI_Tab_Settings.Controls.Add($askme)
$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_Settings)


###############################################################################################

$GUI_Tab_About = New-object System.Windows.Forms.Tabpage
$GUI_Tab_About.UseVisualStyleBackColor = $True 
$GUI_Tab_About.Name = "About" 
$GUI_Tab_About.Text = $text_abouttab


# FANCY ICON
$applogo             = new-object Windows.Forms.PictureBox
$applogo.Width       = 64
$applogo.Height      = 64
$applogo.Image       = $icon
$applogo.Location    = New-Object System.Drawing.Point(128,20)

# Label above input
$abouttitle                     = New-Object System.Windows.Forms.Label
$abouttitle.Size                = New-Object System.Drawing.Size(280,20)
$abouttitle.Left                = ($GUI_Form_MainWindow_leftalign + 85)
$abouttitle.Top                 = 95
$abouttitle.Text                = "-Rocketlaunch!"
$abouttitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 13, [System.Drawing.FontStyle]::Bold)

# Label above input
$aboutsubtitle                     = New-Object System.Windows.Forms.Label
$aboutsubtitle.Size                = New-Object System.Drawing.Size(360,20)
$aboutsubtitle.Left                = ($GUI_Form_MainWindow_leftalign + 40)
$aboutsubtitle.Top                 = 120
$aboutsubtitle.Text                = $text_aboutsubtitle
$aboutsubtitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Italic)

# Label above input
$abouttext                      = New-Object System.Windows.Forms.TextBox
$abouttext.Size                 = New-Object System.Drawing.Size(255,155)
$abouttext.Left                 = ($GUI_Form_MainWindow_leftalign + 25)
$abouttext.Top                  = 150
$abouttext.ReadOnly             = $true
$abouttext.BackColor            = "White"
$abouttext.Multiline            = $true
$abouttext.TextAlign            = "Center"
$abouttext.Text                 = $text_abouttext
$abouttext.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)


[int]$buttonalign = 325

$gotogithub                     = New-Object System.Windows.Forms.Button
$gotogithub.Size                = New-Object System.Drawing.Size (100,25)
$gotogithub.Left                = ($GUI_Form_MainWindow_leftalign)
$gotogithub.Top                 = $buttonalign
$gotogithub.Text                = $text_about_button_repo
$gotogithub.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch"} )

$gotolicense                    = New-Object System.Windows.Forms.Button
$gotolicense.Size               = New-Object System.Drawing.Size (100,25)
$gotolicense.Left               = ($GUI_Form_MainWindow_leftalign + 105)
$gotolicense.Top                = $buttonalign
$gotolicense.Text               = $text_about_button_licence
$gotolicense.Add_Click( {start-process "https://www.gnu.org/licenses/gpl-3.0.html"})


$supportme                      = New-Object System.Windows.Forms.Button
$supportme.Size                 = New-Object System.Drawing.Size (100,25)
$supportme.Left                 = ($GUI_Form_MainWindow_leftalign + 210)
$supportme.Top                  = $buttonalign
$supportme.Text                 = $text_about_button_support
$supportme.Add_Click( {start-process "https://ko-fi.com/teamcons"})


$GUI_Tab_About.Controls.Add($abouttitle)
$GUI_Tab_About.Controls.Add($aboutsubtitle)
$GUI_Tab_About.Controls.Add($applogo)
$GUI_Tab_About.Controls.Add($abouttext)
$GUI_Tab_About.Controls.Add($gotogithub)
$GUI_Tab_About.Controls.Add($supportme)
$GUI_Tab_About.Controls.Add($gotolicense)


$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_About)

