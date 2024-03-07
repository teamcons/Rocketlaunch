﻿



    #==============================================================
    #                                                             =
    #                GUI - Ask the Right Questions                =
    #                                                             =
    #==============================================================






#================
#= INITIAL WORK =

[int]$GUI_Form_MainWindow_leftalign = 15
[int]$GUI_Form_MainWindow_verticalalign = 600


$script:GUI_Form_MainWindow                   = New-Object System.Windows.Forms.Form
$GUI_Form_MainWindow.Text              = -join($APPNAME," - ",$text_aboutsubtitle)
$GUI_Form_MainWindow.Size              = New-Object System.Drawing.Size(775,($GUI_Form_MainWindow_verticalalign + 85 ))
$GUI_Form_MainWindow.MinimumSize       = New-Object System.Drawing.Size(500,170)
$GUI_Form_MainWindow.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)
$GUI_Form_MainWindow.StartPosition     = 'CenterScreen'
$GUI_Form_MainWindow.Topmost           = $True
$GUI_Form_MainWindow.BackColor         = "White"
$GUI_Form_MainWindow.Icon              = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

#==============
#= INPUT TEXT =

$panel_top = New-Object System.Windows.Forms.Panel
$panel_top.Width         = 775
$panel_top.Top           = 0
$panel_top.Height        = 88
$panel_top.Left          = 0
$panel_top.BackColor     = "Orange" #'Ti'
#$panel_top.ForeColor     = "White" #'Ti'
$panel_top.Dock          = "Top"


# FANCY ICON
$pictureBox             = new-object Windows.Forms.PictureBox
$pictureBox.Location    = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,12)
$pictureBox.Anchor      = "Left,Top"
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$pictureBox.Width       = 64
$pictureBox.Height      = 64
$pictureBox.Image       = $img;


# LABEL AND TEXT
# Label above input
$label                  = New-Object System.Windows.Forms.Label
$label.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),12)
$label.Size             = New-Object System.Drawing.Size(300,30)
$label.AutoSize         = $true
$label.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 17, [System.Drawing.FontStyle]::Bold)
$label.Text             = $text_projectname
$label.Anchor           = "Left,Top"

# Input box
$gui_year                  = New-Object System.Windows.Forms.Label
$gui_year.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),52)
$gui_year.Size             = New-Object System.Drawing.Size(20,20)
$gui_year.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)
$gui_year.AutoSize         = $true
$gui_year.Text             = -join($YEAR," -")
$gui_year.Anchor           = "Left,Top"

$script:gui_code                    = New-Object System.Windows.Forms.Combobox
$gui_code.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 130),50)
$gui_code.Size                      = New-Object System.Drawing.Size(210,30)    


# Topmost according to whether checked or not
$gui_keepontop                           = New-Object System.Windows.Forms.Checkbox
#$gui_keepontop.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 365),52)
$gui_keepontop.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 350),52)
$gui_keepontop.Size                      = New-Object System.Drawing.Size(160,20)
$gui_keepontop.Text                      = $text_keepontop
$gui_keepontop.UseVisualStyleBackColor   = $True
$gui_keepontop.Checked                   = $GUI_Form_MainWindow.Topmost
$gui_keepontop.Anchor                    = "Top, Left"
$gui_keepontop.Add_Click({$GUI_Form_MainWindow.Topmost = $gui_keepontop.Checked})
#$panel_top.Controls.Add($gui_keepontop)


$panel_top.controls.add($pictureBox)
$panel_top.Controls.Add($label)
$panel_top.Controls.Add($gui_year)
$panel_top.Controls.Add($gui_code)    
$panel_top.Controls.Add($gui_keepontop)    

$GUI_Form_MainWindow.Controls.Add($panel_top)
$GUI_Form_MainWindow.Add_Shown({$gui_code.Select()})



#===================
#= SOURCE FILES    =

$panel_sourcefile               = New-Object System.Windows.Forms.Panel
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

$script:gui_filesource                 = New-Object System.Windows.Forms.Combobox
$gui_filesource.Location        = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 590),5)
$gui_filesource.Size            = New-Object System.Drawing.Size(140,20)
$gui_filesource.DropDownStyle   = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gui_filesource.Anchor          = "Top,Right"
[void] $gui_filesource.Items.Add($text_from_Outlook) 
[void] $gui_filesource.Items.Add($text_from_Downloads)
[void] $gui_filesource.Items.Add($text_DragNDrop)      
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
$sourcefiles.BorderStyle            = "FixedSingle"

[void]$sourcefiles.Columns.Add($text_columns_Subject,300)
[void]$sourcefiles.Columns.Add($text_columns_Sendername,200)
[void]$sourcefiles.Columns.Add($text_columns_Attachments,70)
[void]$sourcefiles.Columns.Add($text_columns_time,100)

$panel_sourcefile.Controls.Add($labelsourcefiles)
$panel_sourcefile.Controls.Add($gui_filesource)
$panel_sourcefile.Controls.Add($sourcefiles)



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
$templates.BackColor                = "LightBlue"
$templates.GridColor                = "White"
$templates.BorderStyle              = "FixedSingle"

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
    $templates.Columns[$i].Name = -join("0",($i - 1))
    $templates.Columns[$i].Width = $folder_spacing
}


#$templatefile = -join($ScriptPath,"\",$TEMPLATE)
#$templates = load_template $templates $templatefile

[void]$templates.Rows.Add("Minimal","info","orig");
[void]$templates.Rows.Add("Standard TEP","info","orig","trados","to trans","from trans","to proof","from proof","to client")
[void]$templates.Rows.Add("Full TEP","info","orig","trados","to TEP","from TEP","to client")
[void]$templates.Rows.Add("Acolad-MemoQ","info","orig","MemoQ","To client")
[void]$templates.Rows.Add("Proofreading only","info","orig","from proof","to client")
[void]$templates.Rows.Add("Sworn Translation","info","orig","to client")
[void]$templates.Rows.Add("Astrid Special","info","orig","studio","trans","proof","to client")
[void]$templates.Rows.Add("Pizza Margherita","Tomate","Mozzarella","Basilikum","Oliven")


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





