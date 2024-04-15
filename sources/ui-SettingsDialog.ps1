
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
$GUI_Form_MoreStuff.Text                = -join($APPNAME," - ",$text_aboutsubtitle)
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
$moresettingsnota.Top                 = 10
$moresettingsnota.Text                = $text_settingsnota
$moresettingsnota.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)


$CheckIfCreateExplorerQuickAccess                       = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateExplorerQuickAccess.Left                  = $GUI_Form_MainWindow_leftalign
$CheckIfCreateExplorerQuickAccess.Top                   = 35
$CheckIfCreateExplorerQuickAccess.Size                  = New-Object System.Drawing.Size(400,20)
$CheckIfCreateExplorerQuickAccess.Text                  = $text_settings_ExplorerQuickAccess
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

$CheckIfCountWords                                      = New-Object System.Windows.Forms.CheckBox        
$CheckIfCountWords.Left                                 = $GUI_Form_MainWindow_leftalign
$CheckIfCountWords.Top                                  = $CheckIfCreateExplorerQuickAccess.Top + 30
$CheckIfCountWords.Size                                 = New-Object System.Drawing.Size(400,20)
$CheckIfCountWords.Text                                 = $text_settings_Countwords
$CheckIfCountWords.Checked                              = $default_countwords


$CheckIfOpenExplorer                                    = New-Object System.Windows.Forms.CheckBox        
$CheckIfOpenExplorer.Location                           = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,180)
$CheckIfOpenExplorer.Left                               = $GUI_Form_MainWindow_leftalign
$CheckIfOpenExplorer.Top                                = $CheckIfCountWords.Top + 30
$CheckIfOpenExplorer.Size                               = New-Object System.Drawing.Size(400,20)
$CheckIfOpenExplorer.Text                               = $text_settings_OpenExplorer
$CheckIfOpenExplorer.Checked                            = $default_openexplorer

$CheckIfNotify                                          = New-Object System.Windows.Forms.CheckBox        
$CheckIfNotify.Left                                     = $GUI_Form_MainWindow_leftalign
$CheckIfNotify.Top                                      = $CheckIfOpenExplorer.Top + 30
$CheckIfNotify.Size                                     = New-Object System.Drawing.Size(400,20)
$CheckIfNotify.Text                                     = $text_settings_Notify
$CheckIfNotify.Checked                                  = $default_notifywhenfinished


$CheckIfCloseAfter                                      = New-Object System.Windows.Forms.CheckBox        
$CheckIfCloseAfter.Left                                 = $GUI_Form_MainWindow_leftalign
$CheckIfCloseAfter.Top                                  = $CheckIfNotify.Top + 30
$CheckIfCloseAfter.Size                                 = New-Object System.Drawing.Size(400,20)
$CheckIfCloseAfter.Text                                 = $text_settings_CloseAfter
$CheckIfCloseAfter.Checked                              = $default_closeafter




<# # CHANGE LANGUAGE
$label_select_lang                     = New-Object System.Windows.Forms.Label
$label_select_lang.Text                = $text_label_select_lang
$label_select_lang.Top                 = $CheckIfCloseAfter.Top + 35
$label_select_lang.Left                = $GUI_Form_MainWindow_leftalign
$label_select_lang.Size                = New-Object System.Drawing.Size(200,20)

$combobox_select_lang                    = New-Object System.Windows.Forms.Combobox
$combobox_select_lang.Top                = ($label_select_lang.Top - 3)
$combobox_select_lang.Left               = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_lang.Size                = New-Object System.Drawing.Size(100,20)
$combobox_select_lang.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList

[void]$combobox_select_lang.Items.Add($text_lang_german)
[void]$combobox_select_lang.Items.Add($text_lang_french)
[void]$combobox_select_lang.Items.Add($text_lang_spanish)
$combobox_select_lang.SelectedItem = $combobox_select_lang.Items[0] #>


# CHANGE LANGUAGE
$label_select_theme                             = New-Object System.Windows.Forms.Label
$label_select_theme.Text                        = $text_label_select_theme
$label_select_theme.Top                         = $label_select_lang.Top + 35
$label_select_theme.Left                        = $GUI_Form_MainWindow_leftalign
$label_select_theme.Size                        = New-Object System.Drawing.Size(200,20)

$combobox_select_theme                          = New-Object System.Windows.Forms.Combobox
$combobox_select_theme.Top                      = ($label_select_theme.Top - 3)
$combobox_select_theme.Left                     = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_theme.Size                     = New-Object System.Drawing.Size(100,20)
$combobox_select_theme.DropDownStyle            = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$combobox_select_theme.Items.Add("Modern Color")
[void]$combobox_select_theme.Items.Add("Boring")
[void]$combobox_select_theme.Items.Add("Brushed Metal")
[void]$combobox_select_theme.Items.Add("Windows 98")
[void]$combobox_select_theme.Items.Add("Princess Eyebleed")
$combobox_select_theme.SelectedItem = $combobox_select_theme.Items[0]




#$GUI_Tab_Settings.Controls.Add($moresettingstitle)
$GUI_Tab_Settings.Controls.Add($moresettingsnota)
#$GUI_Tab_Settings.Controls.Add($GUI_More_Close)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateExplorerQuickAccess)
#$GUI_Tab_Settings.Controls.Add($CheckIfCreateOutlookFolder)
$GUI_Tab_Settings.Controls.Add($CheckIfExpandArchives)
$GUI_Tab_Settings.Controls.Add($CheckIfCountWords)
$GUI_Tab_Settings.Controls.Add($CheckIfOpenExplorer)
$GUI_Tab_Settings.Controls.Add($CheckIfNotify)
$GUI_Tab_Settings.Controls.Add($CheckIfCloseAfter)

#$GUI_Tab_Settings.Controls.Add($label_select_lang)
#$GUI_Tab_Settings.Controls.Add($combobox_select_lang)
#$GUI_Tab_Settings.Controls.Add($label_select_theme)
#$GUI_Tab_Settings.Controls.Add($combobox_select_theme)


#####################
$helptitle                                              = New-Object System.Windows.Forms.Label
$helptitle.Size                                         = New-Object System.Drawing.Size(280,20)
$helptitle.Left                                         = $GUI_Form_MainWindow_leftalign
$helptitle.Top                                          = 280
$helptitle.Text                                         = $text_settings_helptitle
$helptitle.Font                                         = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$getthedoc                                              = New-Object System.Windows.Forms.Button
$getthedoc.Size                                         = New-Object System.Drawing.Size (150,25)
$getthedoc.Left                                         = $GUI_Form_MainWindow_leftalign
$getthedoc.Top                                          = $buttonalign
$getthedoc.Text                                         = $text_settings_getthedoc
#$getthedoc.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch/raw/main/docs/Manual%20-%20Rocketlaunch.docx"})
$getthedoc.Add_Click( {start-process (-join($ScriptPath,"\documentation\Rocketlaunch Manual.pdf")) } )

<# $askme                                                  = New-Object System.Windows.Forms.Button
$askme.Text                                             = $text_settings_askme
$askme.Size                                             = New-Object System.Drawing.Size (150,25)
$askme.Left                                             = ($GUI_Form_MainWindow_leftalign + 155)
$askme.Top                                              = $getthedoc.Top 
$askme.Add_Click( {start-process "Mailto:stella.menier@gmx.de"}) #>

$GUI_More_Closebutton                               = New-Object System.Windows.Forms.Button
$GUI_More_Closebutton.Text                          = $text_settings_close
$GUI_More_Closebutton.Size                          = New-Object System.Drawing.Size(100,25)
$GUI_More_Closebutton.Left                          = 215
$GUI_More_Closebutton.Top                           = $buttonalign
$GUI_More_Closebutton.Add_Click( {$GUI_Form_MoreStuff.Close() } )

#$GUI_Tab_Settings.Controls.Add($helptitle)
$GUI_Tab_Settings.Controls.Add($getthedoc)
$GUI_Tab_Settings.Controls.Add($GUI_More_Closebutton)

#$GUI_Tab_Settings.Controls.Add($askme)
$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_Settings)


###############################################################################################

$GUI_Tab_About = New-object System.Windows.Forms.Tabpage
$GUI_Tab_About.UseVisualStyleBackColor = $True 
$GUI_Tab_About.Name = "About" 
$GUI_Tab_About.Text = $text_abouttab


# FANCY ICON
$applogo                        = new-object Windows.Forms.PictureBox
$applogo.Width                  = 64
$applogo.Height                 = $applogo.Width
$applogo.Image                  = $image
$applogo.Location               = New-Object System.Drawing.Point(128,20)

#$img                    = (get-item $ScriptPath\assets\icon-mini.ico)
#$pictureBox.Image       = [system.drawing.image]::FromFile($img)

# Label above input
$abouttitle                     = New-Object System.Windows.Forms.Label
$abouttitle.Text                = "-Rocketlaunch!"
$abouttitle.Size                = New-Object System.Drawing.Size(280,20)
$abouttitle.Left                = ($GUI_Form_MainWindow_leftalign + 85)
$abouttitle.Top                 = 95
$abouttitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 13, [System.Drawing.FontStyle]::Bold)

# Label above input
$aboutsubtitle                  = New-Object System.Windows.Forms.Label
$aboutsubtitle.Text             = $text_aboutsubtitle
$aboutsubtitle.Size             = New-Object System.Drawing.Size(360,20)
$aboutsubtitle.Left             = ($GUI_Form_MainWindow_leftalign + 40)
$aboutsubtitle.Top              = 120
$aboutsubtitle.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Italic)

# Label above input
$abouttext                      = New-Object System.Windows.Forms.TextBox
$abouttext.Text                 = $text_abouttext
$abouttext.Size                 = New-Object System.Drawing.Size(255,162)
$abouttext.Left                 = ($GUI_Form_MainWindow_leftalign + 25)
$abouttext.Top                  = 150
$abouttext.ReadOnly             = $true
$abouttext.BackColor            = "White"
$abouttext.Multiline            = $true
$abouttext.TextAlign            = "Center"
$abouttext.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)




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

