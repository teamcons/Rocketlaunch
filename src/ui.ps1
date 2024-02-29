
        #===================================================
        #                GUI - About Dialog                =
        #===================================================


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 



[int]$GUI_Form_MainWindow_leftalign = 5

$global:GUI_Form_MoreStuff                     = New-Object System.Windows.Forms.Form
$GUI_Form_MoreStuff.Text                = $APPNAME
$GUI_Form_MoreStuff.Icon                = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

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

$CheckIfCreateExplorerQuickAccess                = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateExplorerQuickAccess.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,40)
$CheckIfCreateExplorerQuickAccess.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfCreateExplorerQuickAccess.Text           = $text_settings_ExplorerQuickAccess
$CheckIfCreateExplorerQuickAccess.Checked        = $default_createshortcut

$CheckIfCreateOutlookFolder                = New-Object System.Windows.Forms.CheckBox        
$CheckIfCreateOutlookFolder.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,70)
$CheckIfCreateOutlookFolder.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfCreateOutlookFolder.Text           = $text_settings_OutlookFolder
$CheckIfCreateOutlookFolder.Checked        = $default_createoutlookfolder

$CheckIfOpenExplorer                = New-Object System.Windows.Forms.CheckBox        
$CheckIfOpenExplorer.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,100)
$CheckIfOpenExplorer.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfOpenExplorer.Text           = $text_settings_OpenExplorer
$CheckIfOpenExplorer.Checked        = $default_openexplorer

$CheckIfNotify                = New-Object System.Windows.Forms.CheckBox        
$CheckIfNotify.Location       = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,130)
$CheckIfNotify.Size           = New-Object System.Drawing.Size(400,20)
$CheckIfNotify.Text           = $text_settings_Notify
$CheckIfNotify.Checked        = $default_notifywhenfinished

$helptitle                     = New-Object System.Windows.Forms.Label
$helptitle.Size                = New-Object System.Drawing.Size(300,20)
$helptitle.Left                = $GUI_Form_MainWindow_leftalign
$helptitle.Top                 = 180
$helptitle.Text                = $text_settings_helptitle
$helptitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)

$getthedoc                 = New-Object System.Windows.Forms.Button
$getthedoc.Size            = New-Object System.Drawing.Size (180,30)
$getthedoc.Left            = $GUI_Form_MainWindow_leftalign
$getthedoc.Top             = 210
$getthedoc.Text            = $text_settings_getthedoc
$getthedoc.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch/raw/main/docs/Manual%20-%20Rocketlaunch.docx"})

$askme                 = New-Object System.Windows.Forms.Button
$askme.Size            = New-Object System.Drawing.Size (120,30)
$askme.Left            = ($GUI_Form_MainWindow_leftalign + 190)
$askme.Top             = 210
$askme.Text            = $text_settings_askme
$askme.Add_Click( {start-process "Mailto:stella.menier@gmx.de"})


$GUI_More_Close                               = New-Object System.Windows.Forms.Button
$GUI_More_Close.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign ),140)
$GUI_More_Close.Size                          = New-Object System.Drawing.Size(120,25)
$GUI_More_Close.Text                          = $text_OK
$GUI_More_Close.UseVisualStyleBackColor       = $True
#$GUI_More_Close.DialogResult                  = [System.Windows.Forms.DialogResult]::OK
$GUI_Form_MoreStuff.AcceptButton                          = $GUI_More_Close


$GUI_Tab_Settings.Controls.Add($moresettingstitle)
#$GUI_Tab_Settings.Controls.Add($GUI_More_Close)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateExplorerQuickAccess)
$GUI_Tab_Settings.Controls.Add($CheckIfCreateOutlookFolder)
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
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$applogo.Width       = 64
$applogo.Height      = 64
$applogo.Image       = $img;
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


[int]$buttonalign = 320

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











    #==============================================================
    #                                                             =
    #                GUI - Ask the Right Questions                =
    #                                                             =
    #==============================================================






#================
#= INITIAL WORK =

[int]$GUI_Form_MainWindow_leftalign = 15
[int]$GUI_Form_MainWindow_verticalalign = 600


$global:GUI_Form_MainWindow                   = New-Object System.Windows.Forms.Form
$GUI_Form_MainWindow.Text              = $APPNAME
$GUI_Form_MainWindow.Size              = New-Object System.Drawing.Size(775,($GUI_Form_MainWindow_verticalalign + 85 ))
$GUI_Form_MainWindow.MinimumSize       = New-Object System.Drawing.Size(500,180)
#$GUI_Form_MainWindow.MaximumSize       = New-Object System.Drawing.Size(750,550)
#$GUI_Form_MainWindow.AutoSize          = $true
#$GUI_Form_MainWindow.AutoScale         = $true
$GUI_Form_MainWindow.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)
$GUI_Form_MainWindow.StartPosition     = 'CenterScreen'
#$GUI_Form_MainWindow.FormBorderStyle   = 'FixedDialog'
$GUI_Form_MainWindow.Topmost           = $True
$GUI_Form_MainWindow.BackColor         = "White"
$GUI_Form_MainWindow.Icon              = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

#==============
#= INPUT TEXT =

# FANCY ICON
$pictureBox             = new-object Windows.Forms.PictureBox
$pictureBox.Location    = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,15)
$pictureBox.Anchor      = "Left,Top"
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$pictureBox.Width       = 64
$pictureBox.Height      = 64
$pictureBox.Image       = $img;
$GUI_Form_MainWindow.controls.add($pictureBox)

# LABEL AND TEXT
# Label above input
$label                  = New-Object System.Windows.Forms.Label
$label.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),25)
$label.Size             = New-Object System.Drawing.Size(300,30)
$label.AutoSize         = $true
$label.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 14, [System.Drawing.FontStyle]::Bold)
$label.Text             = $text_projectname
$label.Anchor           = "Left,Top"
$GUI_Form_MainWindow.Controls.Add($label)

# Input box
$gui_year                  = New-Object System.Windows.Forms.Label
$gui_year.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),63)
$gui_year.Size             = New-Object System.Drawing.Size(20,20)
$gui_year.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$gui_year.AutoSize         = $true
$gui_year.Text             = -join($YEAR," -")
$gui_year.Anchor           = "Left,Top"
$GUI_Form_MainWindow.Controls.Add($gui_year)


$global:gui_code                    = New-Object System.Windows.Forms.Combobox
$gui_code.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 124),60)
$gui_code.Size                      = New-Object System.Drawing.Size(250,30)    
$GUI_Form_MainWindow.Controls.Add($gui_code)    
$GUI_Form_MainWindow.Add_Shown({$gui_code.Select()})



#===================
#= SOURCE FILES    =

$panel_sourcefile = New-Object System.Windows.Forms.Panel
$panel_sourcefile.Width         = 775
$panel_sourcefile.Top           = 25
$panel_sourcefile.Height        = 200
$panel_sourcefile.Left          = 0
$panel_sourcefile.BackColor     = "White" #'Green'
$panel_sourcefile.Dock          = "Fill"

# Label above input
$labelsourcefiles                  = New-Object System.Windows.Forms.Label
$labelsourcefiles.Location         = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,10)
$labelsourcefiles.Size             = New-Object System.Drawing.Size(240,20)
$labelsourcefiles.Text             = $text_loadfilesfrom
$labelsourcefiles.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)

$gui_filesource                 = New-Object System.Windows.Forms.Combobox
$gui_filesource.Location        = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 590),5)
$gui_filesource.Size            = New-Object System.Drawing.Size(140,20)
$gui_filesource.DropDownStyle   = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gui_filesource.Anchor          = "Top,Right"
[void] $gui_filesource.Items.Add($text_from_Outlook) 
[void] $gui_filesource.Items.Add($text_from_Downloads)   
[void] $gui_filesource.Items.Add($text_nofilesource)  

$gui_filesource.SelectedItem = $default_filesfrom

## Configure the ListView
$sourcefiles                        = New-Object System.Windows.Forms.ListView
$sourcefiles.Location               = New-Object System.Drawing.Size($GUI_Form_MainWindow_leftalign,30) 
$sourcefiles.Size                   = New-Object System.Drawing.Size(730,160) 
$sourcefiles.FullRowSelect          = $True
$sourcefiles.HideSelection          = $false
$sourcefiles.Anchor                 = "Left,Right,Top,Bottom"
$sourcefiles.View                   = [System.Windows.Forms.View]::Details

[void]$sourcefiles.Columns.Add($text_columns_Subject,300)
[void]$sourcefiles.Columns.Add($text_columns_Sendername,200)
[void]$sourcefiles.Columns.Add($text_columns_Attachments,70)
[void]$sourcefiles.Columns.Add($text_columns_time,100)

$panel_sourcefile.Controls.Add($labelsourcefiles)
$panel_sourcefile.Controls.Add($gui_filesource)
$panel_sourcefile.Controls.Add($sourcefiles)
$panel_sourcefile.Show()



#=====================
#= LIST OF TEMPLATES =

$panel_template                         = New-Object System.Windows.Forms.Panel
$panel_template.Width                   = 775
$panel_template.Height                  = 100
$panel_template.Top                     = 260
$panel_template.Left                    = 0
$panel_template.BackColor               = "White" #'Red'
$panel_template.Dock = "Fill"


# Label and button
$labeltemplate                          = New-Object System.Windows.Forms.Label
$labeltemplate.Text                     = $text_usewhichtemplate
$labeltemplate.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$labeltemplate.Left                     = $GUI_Form_MainWindow_leftalign
$labeltemplate.Top                      = 10
$labeltemplate.Size                     = New-Object System.Drawing.Size(300,20)
$labeltemplate.MinimumSize              = New-Object System.Drawing.Size(300,20)
$labeltemplate.MaximumSize              = New-Object System.Drawing.Size(300,20)
$labeltemplate.Anchor                   = "Left,Top"

$gui_browsetemplate                   = New-Object System.Windows.Forms.Button
$gui_browsetemplate.Left              = ($GUI_Form_MainWindow_leftalign + 630)
$gui_browsetemplate.Top               = 5
$gui_browsetemplate.Size              = New-Object System.Drawing.Size(100,25)
$gui_browsetemplate.Text              = $text_loadtemplate
$gui_browsetemplate.Anchor            = "Right,Top"
$gui_browsetemplate.add_click({ 
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
                                                                                InitialDirectory = [Environment]::GetFolderPath('Desktop')
                                                                            }
    $file = $FileBrowser.ShowDialog()    
    $templates = load_template $templates $file
    $templates.Refresh()})

$templates                          = New-Object System.Windows.Forms.DataGridView
$templates.Location                 = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,30)
$templates.Size                     = New-Object System.Drawing.Size(730,70)
$templates.AutoResizeColumns(2)
$templates.Anchor                   = "Left,Right,Top,Bottom"
$templates.BackgroundColor          = "White"
#$templates.GridColor                = "LightBlue"
$templates.GridColor                = "White"
$templates.CellBorderStyle          = "SingleHorizontal"
$templates.SelectionMode            = "FullRowSelect"
$templates.RowHeadersVisible        = $false
$templates.MultiSelect              = $false
$templates.AllowUserToResizeRows    = $false
$templates.ColumnCount = 10
$templates.AutoGenerateColumns = $true

[int]$folder_spacing = 80
$templates.Columns[0].Name = "Vorlage"
$templates.Columns[0].Width = 120
for ($i=1; $i -lt $templates.ColumnCount ; $i++)
{
    $templates.Columns[$i].Name = -join("0",$i)
    $templates.Columns[$i].Width = $folder_spacing
}


#$templatefile = -join($ScriptPath,"\",$TEMPLATE)
#$templates = load_template $templates $templatefile

[void]$templates.Rows.Add("Minimal","info","orig");
[void]$templates.Rows.Add("Standard TEP","info","orig","trados","to trans","from trans","to proof","from proof","to client")
[void]$templates.Rows.Add("Full TEP","info","orig","to TEP","from TEP","to client")
[void]$templates.Rows.Add("Acolad","info","orig","MemoQ","To client")
[void]$templates.Rows.Add("Proofreading only","info","orig","from proof","to client")
[void]$templates.Rows.Add("Sworn Translation","info","orig","to client")


$templates.Rows[0].Selected = $true #.Selected = $true

$panel_template.Controls.Add($labeltemplate)
#$panel_template.Controls.Add($gui_browsetemplate)
$panel_template.Controls.Add($templates)
$panel_template.Show()


$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Anchor                       = "Left,Bottom,Top,Right"
$Split.Top                          = 90
$Split.Height                       = ($GUI_Form_MainWindow_verticalalign - 100 )
$Split.Width                        = 775
$Split.Orientation                  = "Horizontal"
$Split.BackColor                    = "LightBlue"
$Split.SplitterDistance             = 220

$Split.Panel1.Controls.Add($panel_sourcefile)
$Split.Panel2.Controls.Add($panel_template)
$GUI_Form_MainWindow.Controls.Add($Split)


#====================
#= OKCANCEL BUTTONS =

$gui_panel = New-Object System.Windows.Forms.Panel
$gui_panel.Left = 0
$gui_panel.Top = ($GUI_Form_MainWindow_verticalalign)
$gui_panel.Width = 775
$gui_panel.Height = 50
$gui_panel.BackColor = '241,241,241'
$gui_panel.Anchor = "Left,Bottom,Right"

$gui_help                   = New-Object System.Windows.Forms.Button
$gui_help.Location          = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign),10)
$gui_help.Size              = New-Object System.Drawing.Size(120,25)
$gui_help.Text              = $text_help
$gui_help.UseVisualStyleBackColor = $True
$gui_help.Anchor            = "Left, Bottom"
$gui_help.add_click({$GUI_Form_MoreStuff.ShowDialog()})
#[void]$GUI_Form_MainWindow.Controls.Add($gui_help)


# Check if start new trados project
$CheckIfTrados                  = New-Object System.Windows.Forms.CheckBox        
$CheckIfTrados.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 130),12)
$CheckIfTrados.Size             = New-Object System.Drawing.Size(70,20)
$CheckIfTrados.Text             = $text_opentrados
$CheckIfTrados.Checked          = $default_opentrados
$CheckIfTrados.Anchor           = "Top,Left"
#[void]$GUI_Form_MainWindow.Controls.Add($CheckIfTrados)



$gui_okButton                               = New-Object System.Windows.Forms.Button
$gui_okButton.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 480),10)
$gui_okButton.Size                          = New-Object System.Drawing.Size(120,25)
$gui_okButton.Text                          = $text_OK
$gui_okButton.UseVisualStyleBackColor       = $True
$gui_okButton.Anchor                        = "Bottom,Right"
#$gui_okButton.BackColor                     = ”Green”
#$gui_okButton.ForeColor                     = ”White”
$gui_okButton.DialogResult                  = [System.Windows.Forms.DialogResult]::OK
$GUI_Form_MainWindow.AcceptButton                          = $gui_okButton
#[void]$GUI_Form_MainWindow.Controls.Add($gui_okButton)

$gui_cancelButton                           = New-Object System.Windows.Forms.Button
$gui_cancelButton.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 610),10)
$gui_cancelButton.Size                      = New-Object System.Drawing.Size(120,25)
$gui_cancelButton.Text                      = $text_Cancel
$gui_cancelButton.UseVisualStyleBackColor   = $True
$gui_cancelButton.Anchor                    = "Bottom, Right"
#$gui_cancelButton.BackColor                  = ”Red”
#$gui_cancelButton.ForeColor                  = ”White”
$gui_cancelButton.DialogResult              = [System.Windows.Forms.DialogResult]::Cancel
$GUI_Form_MainWindow.CancelButton                          = $gui_cancelButton
#[void]$GUI_Form_MainWindow.Controls.Add($gui_cancelButton)


$gui_panel.Controls.Add($gui_help)
$gui_panel.Controls.Add($CheckIfTrados)
$gui_panel.Controls.Add($gui_okButton)
$gui_panel.Controls.Add($gui_cancelButton)
$gui_panel.Show()

[void]$GUI_Form_MainWindow.Controls.Add($gui_panel)






