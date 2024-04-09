

    #==============================================================
    #                                                             =
    #                GUI - Ask the Right Questions                =
    #                                                             =
    #==============================================================



<#
Define the main interface

#>



#================================
# INITIAL WORK

Write-Output "[START] Loading UI"

[int]$GUI_Form_MainWindow_leftalign = 15
[int]$GUI_Form_MainWindow_verticalalign = 600


$script:GUI_Form_MainWindow                     = New-Object System.Windows.Forms.Form
$GUI_Form_MainWindow.Text                       = -join($APPNAME," - ",$text_aboutsubtitle)
$GUI_Form_MainWindow.Size                       = New-Object System.Drawing.Size(775,($GUI_Form_MainWindow_verticalalign + 85 ))
$GUI_Form_MainWindow.MinimumSize                = New-Object System.Drawing.Size(585,172)
$GUI_Form_MainWindow.StartPosition              = 'CenterScreen'
$GUI_Form_MainWindow.Topmost                    = $default_ontop
$GUI_Form_MainWindow.Icon                       = $icon
$GUI_Form_MainWindow.Add_Closing({Close-All})


#================================
# TOP PANEL

$panel_top = New-Object System.Windows.Forms.Panel
$panel_top.Width                    = 775
$panel_top.Top                      = 0
$panel_top.Height                   = 88
$panel_top.Left                     = 0
$panel_top.Dock                     = "Top"


# FANCY ICON
$pictureBox                         = new-object Windows.Forms.PictureBox
$pictureBox.Location                = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,12)
$pictureBox.Anchor                  = "Left,Top"
$pictureBox.Width                       = 64
$pictureBox.Height                  = 64
$pictureBox.Image                   = $image


# LABEL AND TEXT
# Label above input
$label                              = New-Object System.Windows.Forms.Label
$label.Location                     = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),12)
$label.Size                         = New-Object System.Drawing.Size(300,30)
$label.AutoSize                     = $true
$label.Font                         = New-Object System.Drawing.Font('Calibri', 17, [System.Drawing.FontStyle]::Bold)
$label.Text                         = $text_projectname
$label.Anchor                       = "Left,Top"

# Input box
$gui_year                           = New-Object System.Windows.Forms.Label
$gui_year.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 80),52)
$gui_year.Size                      = New-Object System.Drawing.Size(20,20)
$gui_year.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Regular)
$gui_year.AutoSize                  = $true
$gui_year.Text                      = -join($YEAR," -")
$gui_year.Anchor                    = "Left,Top"

$script:gui_code                    = New-Object System.Windows.Forms.Combobox
$gui_code.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 130),50)
$gui_code.Size                      = New-Object System.Drawing.Size(210,30)




# What outlook folder to create a project folder in ?
$script:gui_folderinoutlook                  = New-Object System.Windows.Forms.Combobox
$gui_folderinoutlook.Location                = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 350),50)
$gui_folderinoutlook.Size                    = New-Object System.Drawing.Size(125,30)
$gui_folderinoutlook.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gui_folderinoutlook.Anchor                  = "Top,Left"

[void]$gui_folderinoutlook.Items.Add("01_QUOTES")
[void]$gui_folderinoutlook.Items.Add("02_CURRENT JOBS")
[void]$gui_folderinoutlook.Items.Add($text_nooutlook)      
$gui_folderinoutlook.SelectedItem = $gui_folderinoutlook.Items[1]

# Check if start new trados project
$CheckIfTrados                              = New-Object System.Windows.Forms.CheckBox        
$CheckIfTrados.Location                     = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 485),50)
$CheckIfTrados.Size                         = New-Object System.Drawing.Size(70,20)
$CheckIfTrados.Text                         = $text_opentrados
$CheckIfTrados.Checked                      = $default_opentrados
$CheckIfTrados.Anchor                       = "Top,Left"

$panel_top.Controls.Add($gui_folderinoutlook)
$panel_top.Controls.Add($CheckIfTrados)



$panel_top.controls.add($pictureBox)
$panel_top.Controls.Add($label)
$panel_top.Controls.Add($gui_year)
$panel_top.Controls.Add($gui_code)    


$GUI_Form_MainWindow.Controls.Add($panel_top)
$GUI_Form_MainWindow.Add_Shown({$gui_code.Select()})



#================================
#= SOURCE FILES

$panel_sourcefile                       = New-Object System.Windows.Forms.Panel
$panel_sourcefile.Width                 = 775
$panel_sourcefile.Top                   = 25
$panel_sourcefile.Height                = 200
$panel_sourcefile.Left                  = 0
$panel_sourcefile.Dock                  = "Fill"

# Label above input
$labelsourcefiles                       = New-Object System.Windows.Forms.Label
$labelsourcefiles.Text                  = $text_label_from_Outlook
$labelsourcefiles.Location              = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,10)
$labelsourcefiles.Size                  = New-Object System.Drawing.Size(450,20)

$sourcefile_refreshButton                               = New-Object System.Windows.Forms.Button
$sourcefile_refreshButton.Text                          = $global:text_sourcefiles_refresh
$sourcefile_refreshButton.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 485),4)
$sourcefile_refreshButton.Size                          = New-Object System.Drawing.Size(95,24)
$sourcefile_refreshButton.Anchor                        = "Top,Right"


$script:gui_filesource                  = New-Object System.Windows.Forms.Combobox
$gui_filesource.Location                = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 590),5)
$gui_filesource.Size                    = New-Object System.Drawing.Size(140,20)
$gui_filesource.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gui_filesource.Anchor                  = "Top,Right"

[void] $gui_filesource.Items.Add($text_from_Outlook) 
[void] $gui_filesource.Items.Add($text_from_Downloads)
[void] $gui_filesource.Items.Add($text_DragNDrop)      
[void] $gui_filesource.Items.Add($text_nofilesource)  
$gui_filesource.SelectedItem = $default_filesfrom

#================================
## Configure the ListView
$sourcefiles                        = New-Object System.Windows.Forms.ListView
$sourcefiles.Location               = New-Object System.Drawing.Size($GUI_Form_MainWindow_leftalign,30) 
$sourcefiles.Size                   = New-Object System.Drawing.Size(730,160) 
$sourcefiles.FullRowSelect          = $True
$sourcefiles.HideSelection          = $false
$sourcefiles.Anchor                 = "Left,Right,Top,Bottom"
$sourcefiles.View                   = [System.Windows.Forms.View]::Details
$sourcefiles.BorderStyle            = "Fixed3D"
$sourcefiles.AllowDrop              = $true    


[void]$sourcefiles.Columns.Add($text_columns_Subject,350)
[void]$sourcefiles.Columns.Add($text_columns_Sendername,200)
[void]$sourcefiles.Columns.Add($text_columns_Attachments,80)
[void]$sourcefiles.Columns.Add($text_columns_time,100)


$panel_sourcefile.Controls.Add($sourcefile_refreshButton)
$panel_sourcefile.Controls.Add($gui_filesource)
$panel_sourcefile.Controls.Add($labelsourcefiles)
$panel_sourcefile.Controls.Add($sourcefiles)



#================================
#= LIST OF TEMPLATES

$panel_template                         = New-Object System.Windows.Forms.Panel
$panel_template.Width                   = 775
$panel_template.Height                  = 100
$panel_template.Top                     = 260
$panel_template.Left                    = 0
$panel_template.Dock = "Fill"


# Label and button
$labeltemplate                          = New-Object System.Windows.Forms.Label
$labeltemplate.Text                     = $text_usewhichtemplate
$labeltemplate.Left                     = $GUI_Form_MainWindow_leftalign
$labeltemplate.Top                      = 10
$labeltemplate.Size                     = New-Object System.Drawing.Size(400,20)
$labeltemplate.Anchor                   = "Left,Top"

$templates_refreshButton                               = New-Object System.Windows.Forms.Button
$templates_refreshButton.Text                          = $global:text_sourcefiles_refresh
$templates_refreshButton.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 485),4)
$templates_refreshButton.Size                          = New-Object System.Drawing.Size(95,24)
$templates_refreshButton.Anchor                        = "Top,Right"
$templates_refreshButton.Add_Click({$templates.Rows.Clear() ; $templates = load_template $templates $templatefile })


# Check if start new trados project
$CheckIfSaveTemplateChanges                  = New-Object System.Windows.Forms.CheckBox        
$CheckIfSaveTemplateChanges.Text             = $text_savetemplatechanges
$CheckIfSaveTemplateChanges.Location         = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 590),8)
$CheckIfSaveTemplateChanges.Size             = New-Object System.Drawing.Size(180,20)
$CheckIfSaveTemplateChanges.Checked          = $default_savetemplatechanges
$CheckIfSaveTemplateChanges.Anchor           = "Top,Right"


$templates                                  = New-Object System.Windows.Forms.DataGridView
$templates.Location                         = New-Object System.Drawing.Point($GUI_Form_MainWindow_leftalign,30)
$templates.Size                             = New-Object System.Drawing.Size(730,70)
$templates.AutoResizeColumns(2)
$templates.Anchor                           = "Left,Right,Top,Bottom"


$templates.GridColor                        = "White"
$templates.BorderStyle                      = "Fixed3D"
$templates.CellBorderStyle                  = "SingleHorizontal"
$templates.SelectionMode                    = "FullRowSelect"
$templates.RowHeadersVisible                = $false
$templates.MultiSelect                      = $false
$templates.AllowUserToResizeRows            = $false
$templates.ColumnHeadersHeightSizeMode      = "DisableResizing"
$templates.ColumnCount                      = 14
$templates.AutoGenerateColumns              = $true

$templates.Columns[0].Name = $text_template_name
$templates.Columns[0].Width = 120

# Fill the grid
for ($i=1; $i -lt $templates.ColumnCount ; $i++)
{
    if ($i -le 10)  {$templates.Columns[$i].Name = -join("0",($i-1)) }
    else            {$templates.Columns[$i].Name = ($i-1)}
    $templates.Columns[$i].Width = 80
}

#$panel_template.Controls.Add($templates_refreshButton)
$panel_template.Controls.Add($CheckIfSaveTemplateChanges)
$panel_template.Controls.Add($labeltemplate)
$panel_template.Controls.Add($templates)
#$panel_template.Show()


$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Anchor                       = "Left,Bottom,Top,Right"
$Split.Top                          = 90
$Split.Height                       = ($GUI_Form_MainWindow_verticalalign - 100 )
$Split.Width                        = 775
$Split.Orientation                  = "Horizontal"
$Split.SplitterDistance             = 220

$Split.Panel1.Controls.Add($panel_sourcefile)
$Split.Panel2.Controls.Add($panel_template)
$GUI_Form_MainWindow.Controls.Add($Split)


#====================
#= BOTTOM PANEL

$bottom_panel                   = New-Object System.Windows.Forms.Panel
$bottom_panel.Left              = 0
$bottom_panel.Top               = $GUI_Form_MainWindow_verticalalign
$bottom_panel.Width             = 775
$bottom_panel.Height            = 50
$bottom_panel.Anchor            = "Left,Bottom,Right"

# This one offers more settings
$gui_help                       = New-Object System.Windows.Forms.Button
$gui_help.Location              = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign),10)
$gui_help.Size                  = New-Object System.Drawing.Size(120,25)
$gui_help.Text                  = $text_help
$gui_help.Anchor                = "Left, Bottom"

# Topmost according to whether checked or not
$gui_keepontop                           = New-Object System.Windows.Forms.Checkbox
#$gui_keepontop.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 365),52)
$gui_keepontop.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 130),13)
$gui_keepontop.Size                      = New-Object System.Drawing.Size(160,20)
$gui_keepontop.Text                      = $text_keepontop
$gui_keepontop.Checked                   = $GUI_Form_MainWindow.Topmost
$gui_keepontop.Anchor                    = "Top, Left"






# This one creates the project
$gui_okButton                               = New-Object System.Windows.Forms.Button
$gui_okButton.Location                      = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 480),10)
$gui_okButton.Size                          = New-Object System.Drawing.Size(120,25)
$gui_okButton.Text                          = $text_OK
$gui_okButton.Anchor                        = "Bottom,Right"



$gui_cancelButton                           = New-Object System.Windows.Forms.Button
$gui_cancelButton.Location                  = New-Object System.Drawing.Point(($GUI_Form_MainWindow_leftalign + 610),10)
$gui_cancelButton.Size                      = New-Object System.Drawing.Size(120,25)
$gui_cancelButton.Text                      = $text_Cancel
$gui_cancelButton.Anchor                    = "Bottom, Right"


$gui_help.UseVisualStyleBackColor           = $True
$gui_okButton.UseVisualStyleBackColor       = $True
$gui_cancelButton.UseVisualStyleBackColor   = $True




$bottom_panel.Controls.Add($gui_help)
$bottom_panel.Controls.Add($gui_okButton)
$bottom_panel.Controls.Add($gui_cancelButton)


$bottom_panel.Controls.Add($gui_keepontop)    

[void]$GUI_Form_MainWindow.Controls.Add($bottom_panel)






