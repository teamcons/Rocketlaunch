
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

####################################

$GUI_Tab_Settings = New-object System.Windows.Forms.Tabpage
$GUI_Tab_Settings.Name = "Advanced" 
$GUI_Tab_Settings.Text = $text.Settings.settingstag
$GUI_Tab_Settings.UseVisualStyleBackColor = $True 

# Label above input
$moresettingstitle                     = New-Object System.Windows.Forms.Label
$moresettingstitle.Size                = New-Object System.Drawing.Size(300,20)
$moresettingstitle.Left                = $GUI_Form_MainWindow_leftalign
$moresettingstitle.Top                 = 10
$moresettingstitle.Text                = $text.Settings.settingstag
$moresettingstitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$moresettingsnota                     = New-Object System.Windows.Forms.Label
$moresettingsnota.Size                = New-Object System.Drawing.Size(300,20)
$moresettingsnota.Left                = $GUI_Form_MainWindow_leftalign
$moresettingsnota.Top                 = 10
$moresettingsnota.Text                = $text.Settings.settingsnota
$moresettingsnota.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)


$CheckIfCreateExplorerQuickAccess                       = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateExplorerQuickAccess.Left                  = $GUI_Form_MainWindow_leftalign
$CheckIfCreateExplorerQuickAccess.Top                   = 35
$CheckIfCreateExplorerQuickAccess.Size                  = New-Object System.Drawing.Size(400,20)
$CheckIfCreateExplorerQuickAccess.Text                  = $text.Settings.ExplorerQuickAccess
$CheckIfCreateExplorerQuickAccess.Checked               = $default_createshortcut

<# $CheckIfCreateOutlookFolder                             = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateOutlookFolder.Left                        = $GUI_Form_MainWindow_leftalign
$CheckIfCreateOutlookFolder.Top                         = $CheckIfCreateExplorerQuickAccess.Top + 30
$CheckIfCreateOutlookFolder.Size                        = New-Object System.Drawing.Size(400,20)
$CheckIfCreateOutlookFolder.Text                        = $text_settings_OutlookFolder
$CheckIfCreateOutlookFolder.Checked                     = $default_createoutlookfolder #>

<# 
$CheckIfExpandArchives                = New-Object System.Windows.Forms.CheckBox        
$CheckIfExpandArchives.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,120)
$CheckIfExpandArchives.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfExpandArchives.Text           = $text_settings_ExpandArchives
$CheckIfExpandArchives.Checked        = $default_expandarchives
 #>

<# $CheckIfCountWords                                      = New-Object System.Windows.Forms.CheckBox        
$CheckIfCountWords.Left                                 = $GUI_Form_MainWindow_leftalign
$CheckIfCountWords.Top                                  = $CheckIfCreateExplorerQuickAccess.Top + 30
$CheckIfCountWords.Size                                 = New-Object System.Drawing.Size(400,20)
$CheckIfCountWords.Text                                 = $text_settings_Countwords
$CheckIfCountWords.Checked                              = $default_countwords #>


$CheckIfOpenExplorer                                    = New-Object System.Windows.Forms.CheckBox        
$CheckIfOpenExplorer.Location                           = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,180)
$CheckIfOpenExplorer.Left                               = $GUI_Form_MainWindow_leftalign
$CheckIfOpenExplorer.Top                                = $CheckIfCreateExplorerQuickAccess.Top + 30
$CheckIfOpenExplorer.Size                               = New-Object System.Drawing.Size(400,20)
$CheckIfOpenExplorer.Text                               = $text.Settings.OpenExplorer
$CheckIfOpenExplorer.Checked                            = $default_openexplorer

$CheckIfArchiveFolder                                          = New-Object System.Windows.Forms.CheckBox        
$CheckIfArchiveFolder.Left                                     = $GUI_Form_MainWindow_leftalign
$CheckIfArchiveFolder.Top                                      = $CheckIfOpenExplorer.Top + 30
$CheckIfArchiveFolder.Size                                     = New-Object System.Drawing.Size(400,20)
$CheckIfArchiveFolder.Text                                     = $text.Settings.createarchivefolder
$CheckIfArchiveFolder.Checked                                  = $default_createarchivefolder


$CheckIfNotify                                          = New-Object System.Windows.Forms.CheckBox        
$CheckIfNotify.Left                                     = $GUI_Form_MainWindow_leftalign
$CheckIfNotify.Top                                      = $CheckIfArchiveFolder.Top + 30
$CheckIfNotify.Size                                     = New-Object System.Drawing.Size(400,20)
$CheckIfNotify.Text                                     = $text.Settings.Notify
$CheckIfNotify.Checked                                  = $default_notifywhenfinished


$CheckIfCloseAfter                                      = New-Object System.Windows.Forms.CheckBox        
$CheckIfCloseAfter.Left                                 = $GUI_Form_MainWindow_leftalign
$CheckIfCloseAfter.Top                                  = $CheckIfNotify.Top + 30
$CheckIfCloseAfter.Size                                 = New-Object System.Drawing.Size(400,20)
$CheckIfCloseAfter.Text                                 = $text.Settings.CloseAfter
$CheckIfCloseAfter.Checked                              = $default_closeafter



#$GUI_Tab_Settings.Controls.Add($moresettingstitle)
$GUI_Tab_Settings.Controls.Add($moresettingsnota)
#$GUI_Tab_Settings.Controls.Add($GUI_More_Close)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateExplorerQuickAccess)
#$GUI_Tab_Settings.Controls.Add($CheckIfCreateOutlookFolder)
$GUI_Tab_Settings.Controls.Add($CheckIfExpandArchives)
$GUI_Tab_Settings.Controls.Add($CheckIfOpenExplorer)
$GUI_Tab_Settings.Controls.Add($CheckIfArchiveFolder)   
$GUI_Tab_Settings.Controls.Add($CheckIfNotify)
$GUI_Tab_Settings.Controls.Add($CheckIfCloseAfter)



#####################
$helptitle                                              = New-Object System.Windows.Forms.Label
$helptitle.Size                                         = New-Object System.Drawing.Size(280,20)
$helptitle.Left                                         = $GUI_Form_MainWindow_leftalign
$helptitle.Top                                          = 280
$helptitle.Text                                         = $text.Settings.helptitle
$helptitle.Font                                         = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$getthedoc                                              = New-Object System.Windows.Forms.Button
$getthedoc.Size                                         = New-Object System.Drawing.Size (205,25)
$getthedoc.Left                                         = $GUI_Form_MainWindow_leftalign
$getthedoc.Top                                          = $buttonalign
$getthedoc.Text                                         = $text.Settings.getthedoc
#$getthedoc.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch/raw/main/docs/Manual%20-%20Rocketlaunch.docx"})
$getthedoc.Add_Click( {start-process (-join($ScriptPath,"\documentation\Rocketlaunch Manual.pdf")) } )


$GUI_More_Closebutton                               = New-Object System.Windows.Forms.Button
$GUI_More_Closebutton.Text                          = $text.Settings.close
$GUI_More_Closebutton.Size                          = New-Object System.Drawing.Size(100,25)
$GUI_More_Closebutton.Left                          = 215
$GUI_More_Closebutton.Top                           = $buttonalign
$GUI_More_Closebutton.Add_Click( {$GUI_Form_MoreStuff.Close() } )

#$GUI_Tab_Settings.Controls.Add($helptitle)
$GUI_Tab_Settings.Controls.Add($getthedoc)
$GUI_Tab_Settings.Controls.Add($GUI_More_Closebutton)

#$GUI_Tab_Settings.Controls.Add($askme)
$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_Settings)