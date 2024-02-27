#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1"
# Usage: Compiled form, just double-click.
#
#
#----------------STEPS----------------
#
# Initialization
# GUI
# Processing Input
# Build the project
# Bonus
#
#-------------------------------------


#===============================================
#                Initialization                =
#===============================================


#========================================
# Fancy !
Write-Output "================================"
Write-Output "=        -ROCKETLAUNCH!        ="
Write-Output "================================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Start new projects really, really quick!"
Write-Output "CC0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output ""
Write-Output ""


#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$APPNAME            = "-Rocketlaunch!"

# Load templates from a csv in same place as executable
#[string]$LOAD_TEMPLATES_FROM = $MyInvocation.MyCommand.Path
[string]$ROOTSTRUCTURE      = "M:\9_JOBS_XTRF\"
[regex]$CODEPATTERN         = -join($YEAR,"-[0-9]")
[string]$YEAR               = get-date –f yyyy

[string]$TEMPLATE               = "vorlagen.csv"
[string]$TEMPLATEDELIMITER               = ";"


if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
{ 
   $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
}
else
{ 
   $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
   if (!$ScriptPath){ $ScriptPath = "." } 
}


#========================================
# Localization
Import-Module $ScriptPath/text.ps1
Import-Module $ScriptPath/defaults.ps1

init_text
init_defaults

Import-Module $ScriptPath/internals.ps1
init_outlook_backend


#==========================================
# Try to predict what next number would be 
# Catch: have at least first part of code


    Write-Output "[STARTUP] Dircode prediction"
    try
    {
        # 
        Set-Location $ROOTSTRUCTURE
        Set-Location (Get-ChildItem 2024_* -Directory | Select-Object -Last 1)   
        $PREDICT_CODE                =  (Get-ChildItem -Directory | Select-Object -Last 1).Name.Substring(5,4)
        [int]$PREDICT_CODE                =  [int]$PREDICT_CODE + 1
        [bool]$CODE_PREDICTED       = $true
        #[string]$PREDICT_CODE   =  -join($YEAR,"-",$PREDICT_CODE,"_")
        Write-Output "[PREDICTED] Next is $PREDICT_CODE"
    }
    catch
    {
        [bool]$CODE_PREDICTED       = $false
        #$PREDICT_CODE = -join($YEAR,"-")
    }


#==============================================================================================================================================================================
#                GUI - About Dialog                =
#===================================================



#================
#= INITIAL WORK =


# Imports
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

Import-Module $ScriptPath/ui.ps1
$GUI_Form_MoreStuff = build_about_gui




#==============================================================
#                GUI - Ask the Right Questions                =
#==============================================================



#================
#= INITIAL WORK =


# Imports
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 


[int]$form_leftalign = 15
[int]$form_verticalalign = 600


$form                   = New-Object System.Windows.Forms.Form
$form.Text              = $APPNAME
$form.Size              = New-Object System.Drawing.Size(775,($form_verticalalign + 85 ))
$form.MinimumSize       = New-Object System.Drawing.Size(500,180)
#$form.MaximumSize       = New-Object System.Drawing.Size(750,550)
#$form.AutoSize          = $true
#$form.AutoScale         = $true
$form.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)
$form.StartPosition     = 'CenterScreen'
#$form.FormBorderStyle   = 'FixedDialog'
$form.Topmost           = $True
$form.BackColor         = "White"
$form.Icon              = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

#==============
#= INPUT TEXT =

# FANCY ICON
$pictureBox             = new-object Windows.Forms.PictureBox
$pictureBox.Location    = New-Object System.Drawing.Point($form_leftalign,15)
$pictureBox.Anchor      = "Left,Top"
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$pictureBox.Width       = 64
$pictureBox.Height      = 64
$pictureBox.Image       = $img;
$form.controls.add($pictureBox)

# LABEL AND TEXT
# Label above input
$label                  = New-Object System.Windows.Forms.Label
$label.Location         = New-Object System.Drawing.Point(($form_leftalign + 80),25)
$label.Size             = New-Object System.Drawing.Size(300,30)
$label.AutoSize         = $true
$label.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 14, [System.Drawing.FontStyle]::Bold)
$label.Text             = $text_projectname
$label.Anchor           = "Left,Top"
$form.Controls.Add($label)

# Input box
$gui_year                  = New-Object System.Windows.Forms.Label
$gui_year.Location         = New-Object System.Drawing.Point(($form_leftalign + 80),63)
$gui_year.Size             = New-Object System.Drawing.Size(20,20)
$gui_year.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)
$gui_year.AutoSize         = $true
$gui_year.Text             = -join($YEAR," -")
$gui_year.Anchor           = "Left,Top"
$form.Controls.Add($gui_year)

# If we have a predicted code we have a numerical value and can offer next codes
if ($CODE_PREDICTED -eq $true)
{
    $gui_code                 = New-Object System.Windows.Forms.Combobox
    $gui_code.Location       = New-Object System.Drawing.Point(($form_leftalign + 124),60)
    $gui_code.Size           = New-Object System.Drawing.Size(170,30)    
    [void] $gui_code.Items.Add( -join($PREDICT_CODE,"_") )  
    [void] $gui_code.Items.Add( -join(($PREDICT_CODE + 1),"_") )  
    [void] $gui_code.Items.Add( -join(($PREDICT_CODE + 2),"_") )  
    [void] $gui_code.Items.Add( -join(($PREDICT_CODE + 3),"_") )  
    $gui_code.SelectedItem = $gui_code.Items[0]
    $form.Controls.Add($gui_code)    
    $form.Add_Shown({$gui_code.Select()})
}
else
{
    $gui_code                = New-Object System.Windows.Forms.TextBox
    $gui_code.Location       = New-Object System.Drawing.Point(($form_leftalign + 123),60)
    $gui_code.Size           = New-Object System.Drawing.Size(170,30)
    $gui_code.Text           = ""
    $form.Controls.Add($gui_code)
    $form.Add_Shown({$gui_code.Select()})
}


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
$labelsourcefiles.Location         = New-Object System.Drawing.Point($form_leftalign,10)
$labelsourcefiles.Size             = New-Object System.Drawing.Size(240,20)
$labelsourcefiles.Text             = $text_loadfilesfrom
$labelsourcefiles.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Regular)

$gui_filesource                 = New-Object System.Windows.Forms.Combobox
$gui_filesource.Location        = New-Object System.Drawing.Point(($form_leftalign + 590),5)
$gui_filesource.Size            = New-Object System.Drawing.Size(140,20)
$gui_filesource.DropDownStyle   = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gui_filesource.Anchor          = "Top,Right"
[void] $gui_filesource.Items.Add($text_from_Outlook) 
[void] $gui_filesource.Items.Add($text_from_Downloads)   
[void] $gui_filesource.Items.Add($text_nofilesource)  

$gui_filesource.SelectedItem = $default_filesfrom

## Configure the ListView
$sourcefiles                        = New-Object System.Windows.Forms.ListView
$sourcefiles.Location               = New-Object System.Drawing.Size($form_leftalign,30) 
$sourcefiles.Size                   = New-Object System.Drawing.Size(730,160) 
$sourcefiles.FullRowSelect          = $True
$sourcefiles.HideSelection          = $false
$sourcefiles.Anchor                 = "Left,Right,Top,Bottom"
$sourcefiles.View                   = [System.Windows.Forms.View]::Details

[void]$sourcefiles.Columns.Add($text_columns_Subject,300)
[void]$sourcefiles.Columns.Add($text_columns_Sendername,200)
[void]$sourcefiles.Columns.Add($text_columns_Attachments,70)
[void]$sourcefiles.Columns.Add($text_columns_time,100)


# Look for emails with attachments
# For each email, we look attachments and count the ones with supported formats
# We are not interested in junk like image001.jpg etc... which is signatures and stuff

#workflow superfast
#{


$allgoodmails = New-Object -TypeName 'System.Collections.ArrayList'
foreach ($mail in $allmails)
{
    [bool]$AddToGoodMails = $false
    [int]$CountGoodAttachments = 0
    foreach ( $attach in $mail.Attachments ) 
    {
        #echo $attach.FileName
        if ($attach.FileName -match  ".(pdf|doc|docx|xls|xlsx|ppt|pptx|xml|idml|csv|txt|zip)" )
        {
            echo (-join("MATCH:",$attach.FileName))
            $AddToGoodMails = $true
            [int]$CountGoodAttachments += 1

        }   # End of each mails
        #else {echo (-join("NOTMATCH:",$attach.FileName)) }  
    } #End of checking attachments

    # we found one with attachment !
    if ($AddToGoodMails -eq $true)
    {
        # Currently observed one is a good one
        $allgoodmails.Add($mail)

        # Add to da list
        $sourcefilesItem = New-Object System.Windows.Forms.ListViewItem($mail.Subject)
        [void]$sourcefilesItem.Subitems.Add($mail.SenderName)
        [void]$sourcefilesItem.Subitems.Add($CountGoodAttachments)
        [void]$sourcefilesItem.Subitems.Add($mail.ReceivedTime.ToString("HH:mm"))
        [void]$sourcefiles.Items.Add($sourcefilesItem)
        $goodmailindex += 1
    } # End of adding goodmail

} # End of looking for emails with attachments


#}


# Add the ListView to the Form
try { 
    $sourcefiles.Items[0].Selected = $true 
}
catch {
    Write-Output "No mail with relevant attach !"
}

try { 
    $allgoodmails.Item(0).SenderEmailAddress -match "@(?<content>.*).com"
    
    #$attempt_at_companyname         = $matches["content"]
    $attempt_at_companyname         = [cultureinfo]::GetCultureInfo("de-DE").TextInfo.ToTitleCase($attempt_at_companyname)
    echo $attempt_at_companyname
    [string]$gui_code.Items[0] = -join($PREDICT_CODE,"_",$attempt_at_companyname )
    [string]$gui_code.Items[1] = -join(($PREDICT_CODE + 1),"_",$attempt_at_companyname )  
    [string]$gui_code.Items[2] = -join(($PREDICT_CODE + 2),"_",$attempt_at_companyname )  
    [string]$gui_code.Items[3] = -join(($PREDICT_CODE + 3),"_",$attempt_at_companyname )
}
catch {
    Write-Output "Messy email !"
}





$panel_sourcefile.Controls.Add($labelsourcefiles)
$panel_sourcefile.Controls.Add($gui_filesource)
$panel_sourcefile.Controls.Add($sourcefiles)
$panel_sourcefile.Show()

#$form.Controls.Add($labelsourcefiles)
#$form.Controls.Add($gui_filesource)
#$form.Controls.Add($sourcefiles)




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
$labeltemplate.Left                     = $form_leftalign
$labeltemplate.Top                      = 10
$labeltemplate.Size                     = New-Object System.Drawing.Size(300,20)
$labeltemplate.MinimumSize              = New-Object System.Drawing.Size(300,20)
$labeltemplate.MaximumSize              = New-Object System.Drawing.Size(300,20)
$labeltemplate.Anchor                   = "Left,Top"



$gui_browsetemplate                   = New-Object System.Windows.Forms.Button
$gui_browsetemplate.Left              = ($form_leftalign + 630)
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
$templates.Location                 = New-Object System.Drawing.Point($form_leftalign,30)
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

$templates.Rows.Add("Minimal","info","orig");
$templates.Rows.Add("Standard TEP","info","orig","trados","to trans","from trans","to proof","from proof","to client")
$templates.Rows.Add("Full TEP","info","orig","to TEP","from TEP","To client")
$templates.Rows.Add("Acolad","info","orig","MemoQ","To client")

$templates.Rows[0].Selected = $true #.Selected = $true

$panel_template.Controls.Add($labeltemplate)
#$panel_template.Controls.Add($gui_browsetemplate)
$panel_template.Controls.Add($templates)
$panel_template.Show()


$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Anchor                       = "Left,Bottom,Top,Right"
$Split.Top                          = 90
$Split.Height                       = ($form_verticalalign - 100 )
$Split.Width                        = 775
$Split.Orientation                  = "Horizontal"
$Split.BackColor                    = "LightBlue"
$Split.SplitterDistance             = 190

$Split.Panel1.Controls.Add($panel_sourcefile)
$Split.Panel2.Controls.Add($panel_template)
$form.Controls.Add($Split)


#====================
#= OKCANCEL BUTTONS =

$gui_panel = New-Object System.Windows.Forms.Panel
$gui_panel.Left = 0
$gui_panel.Top = ($form_verticalalign)
$gui_panel.Width = 775
$gui_panel.Height = 50
$gui_panel.BackColor = '241,241,241'
$gui_panel.Anchor = "Left,Bottom,Right"

$gui_help                   = New-Object System.Windows.Forms.Button
$gui_help.Location          = New-Object System.Drawing.Point(($form_leftalign),10)
$gui_help.Size              = New-Object System.Drawing.Size(120,25)
$gui_help.Text              = $text_help
$gui_help.UseVisualStyleBackColor = $True
$gui_help.Anchor            = "Left, Bottom"
$gui_help.add_click({$GUI_Form_MoreStuff.ShowDialog()})
#[void]$form.Controls.Add($gui_help)


# Check if start new trados project
$CheckIfTrados                  = New-Object System.Windows.Forms.CheckBox        
$CheckIfTrados.Location         = New-Object System.Drawing.Point(($form_leftalign + 130),12)
$CheckIfTrados.Size             = New-Object System.Drawing.Size(70,20)
$CheckIfTrados.Text             = $text_opentrados
$CheckIfTrados.Checked          = $default_opentrados
$CheckIfTrados.Anchor           = "Top,Left"
#[void]$form.Controls.Add($CheckIfTrados)



$gui_okButton                               = New-Object System.Windows.Forms.Button
$gui_okButton.Location                      = New-Object System.Drawing.Point(($form_leftalign + 480),10)
$gui_okButton.Size                          = New-Object System.Drawing.Size(120,25)
$gui_okButton.Text                          = $text_OK
$gui_okButton.UseVisualStyleBackColor       = $True
$gui_okButton.Anchor                        = "Bottom,Right"
#$gui_okButton.BackColor                     = ”Green”
#$gui_okButton.ForeColor                     = ”White”
$gui_okButton.DialogResult                  = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton                          = $gui_okButton
#[void]$form.Controls.Add($gui_okButton)

$gui_cancelButton                           = New-Object System.Windows.Forms.Button
$gui_cancelButton.Location                  = New-Object System.Drawing.Point(($form_leftalign + 610),10)
$gui_cancelButton.Size                      = New-Object System.Drawing.Size(120,25)
$gui_cancelButton.Text                      = $text_Cancel
$gui_cancelButton.UseVisualStyleBackColor   = $True
$gui_cancelButton.Anchor                    = "Bottom, Right"
#$gui_cancelButton.BackColor                  = ”Red”
#$gui_cancelButton.ForeColor                  = ”White”
$gui_cancelButton.DialogResult              = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton                          = $gui_cancelButton
#[void]$form.Controls.Add($gui_cancelButton)


$gui_panel.Controls.Add($gui_help)
$gui_panel.Controls.Add($CheckIfTrados)
$gui_panel.Controls.Add($gui_okButton)
$gui_panel.Controls.Add($gui_cancelButton)
$gui_panel.Show()

[void]$form.Controls.Add($gui_panel)



#==============
#= WRAP IT UP =


$result = $form.ShowDialog()

[string]$PROJECTNAME        = $gui_code.Text 
$PROJECTTEMPLATE            = $templates.SelectedItems.Text

# Cancel culture
# Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Output "[INPUT] Got Cancel. Aw. Exit."
        exit
    }
Write-Output "[INPUT] Got: $PROJECTNAME"


#====================================================================================================================================================================
#                     Processing Le input                     =
#==============================================================


# Empty, so go on with what was initially predicted
if ("$PROJECTNAME" -notmatch "^[0-9]" )
{

    $PROJECTNAME = -join($PREDICT_CODE,$PROJECTNAME)
    Write-Output "Its words. Now: $PROJECTNAME"
}

# Remove invalid character, just in case
$PROJECTNAME = $PROJECTNAME.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
Write-Output "Removed invalid. Now: $PROJECTNAME"


# is it missing zeros
if ($PROJECTNAME -match "^[0-9][0-9][0-9]")
{
    $PROJECTNAME = -join("0",$PROJECTNAME)
    Write-Output "Missing first zero. Now: $PROJECTNAME"

}
elseif ($PROJECTNAME -match "^[0-9][0-9]")
{
    $PROJECTNAME = -join("00",$PROJECTNAME)
    Write-Output "Missing two zero. Now: $PROJECTNAME"
}
elseif ($PROJECTNAME -match "^[0-9]")
{
    $PROJECTNAME = -join("000",$PROJECTNAME)
    Write-Output "Missing three zero. Now: $PROJECTNAME"
}

# Add year
$PROJECTNAME = -join($YEAR,"-",$PROJECTNAME)




##### Ultimate check
try { $DIRCODE = $PROJECTNAME.SubString(0, 9) }
catch {
	$ERRORTEXT="Projektcode ist unpassend !!!
Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
Angegeben: $PROJECTCODE"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)

exit }



#==============================================================
#                      Build The Project                      =
#==============================================================
# REBUILT THE WHOLE TREE
# Its in... year, underscore
$BASEFOLDER     = -join($ROOTSTRUCTURE,$DIRCODE.Substring(0,4),"_")
$BASEFOLDER     = -join($BASEFOLDER,$DIRCODE.Substring(5,2),"00-",$DIRCODE.Substring(5,2),"99")

# If the folder with project numbers in range do not exist, just create it lol
if (!(Test-Path $BASEFOLDER -PathType Container)) {
    Write-Output "[CREATE] Range folder in tree: $BASEFOLDER"
    New-Item -ItemType Directory -Force -Path "$BASEFOLDER"
}
$BASEFOLDER = -join($BASEFOLDER,"\",$PROJECTNAME)



Write-Output "[CREATE] Base folder: $BASEFOLDER"
New-Item -ItemType Directory -Path "$BASEFOLDER"

# Count folder number
[int]$foldernumber = 0 

# CREATE ALLLLL THE FOLDERS
# Skip the first element cuz no
foreach ($folder in ($templates.Rows[$templates.CurrentCell.RowIndex].Cells | Select-Object -Skip 1 ) )
{
    echo $folder
    if ($folder.Value )
    {
        #Append folder number at start, construct full path
        [string]$newfolder = -join("0",$foldernumber,"_",$folder.Value)
        [string]$newfolder = -join($BASEFOLDER,'\',$newfolder)

        # Say what we do, do it
        Write-Output "[CREATE] folder: $newfolder"
        New-Item -ItemType Directory -Path $newfolder

        # Next folder get next number
        [int]$foldernumber = $foldernumber + 1 
    }
}




# PIN TO EXPLORER
if ($CheckIfCreateExplorerQuickAccess.Checked)
{
    Write-Output "[CREATE] Shortcut in File explorer"
    $o = new-object -com shell.application
    $o.Namespace($BASEFOLDER).Self.InvokeVerb("pintohome")
}

if ($CheckIfCreateOutlookFolder.Checked)
{
    Write-Output "[CREATE] Folder in Outlook"
    [void]$ns.Folders.Item(1).Folders.Item("Posteingang").Folders.Item("02_ONGOING JOBS").Folders.Add($PROJECTNAME)
}









#======================================================================================================================================================================
#                      POSTPROCESSING                      =
#===========================================================



#==========================
#= INCLUDE ORIGINAL FILES =


# If user asked to include source files, include those in new folder, with naming conventions
if ($gui_filesource.SelectedItems.Text -notmatch $text_nofilesource)
{

    # CHECK WE HAVE THE MINIMUM FOLDERS
    # BECAUSE WE DONT KNOW WHAT TEMPLATE USER USED
    # IF THE STANDARD MINIMUM ISNT THERE, JUST USE BASE FOLDER INSTEAD
    [string]$INFO = "$BASEFOLDER\00_info"
    [string]$ORIG = "$BASEFOLDER\01_orig"

    if ($gui_filesource.SelectedItem -match $text_from_Outlook )
    {
        Write-Output "[DETECTED] Get source files from email"
        $sourcemail = $allgoodmails[$sourcefiles.SelectedItems.Index]
        foreach ($attachment in $sourcemail.Attachments)
        {
            if ($attachment.FileName -notmatch "^image[0-9][0-9][0-9]")
            {
                Write-Output (-join($ORIG,"\",$attachment.FileName))
                $attachment.SaveAsFile( -join($ORIG,"\",$attachment.FileName) )

            }
        } # End of attachment processing




    } # End of process outlook inclusion
    elseif ($gui_filesource.SelectedItem -match $text_from_Downloads )
    {
         Write-Output "[DETECTED] Load source files"
         # Grab source files
         $load_files = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
             InitialDirectory    = $default_fromdisk
             Multiselect         = $true
             Title               = $APPNAME
         }
         $null = $load_files.ShowDialog()
         Write-Output "[INPUT] Got:"
         Write-Output $SOURCEFILES.FileNames

         foreach ($file in $load_files)
         {
             Write-Output "Moving $file"
             Move-Item -Path $file -Destination $ORIG
         }
    } # End of user load themselves


    # Before processing each source file,
    # Deal with the archives first
    Write-Output "Extracting all archives..."
    Get-ChildItem -Path $ORIG -Filter *.zip | Expand-Archive -DestinationPath $ORIG

    # PROCESS EACH SOURCE FILE
    # Rename and move file
    # Add count to total count and CSV
    # Ignore structure folders
    foreach ($file in (Get-ChildItem -Path "$ORIG" -Exclude "^[0-9][0-9]_" ))
    {
        echo "$ORIG"
        $newname = -join($DIRCODE,"_",$file.BaseName,"_orig",$file.Extension)
        Write-Output "[RENAME] As $newname"
        Rename-Item -Path $file.FullName -Newname "$newname"

        
        <# # ONLY IF ANALYSIS WISHED
        if ($CheckIfAnalysis.CheckState.ToString() -eq "Checked")
        {
            # Use different backend depending on what needed
            # Each time, check the extension to know what we deal with
            if ("$newname" -match ".[doc|docx]" )
            {
                # OPEN IN WORD, PROCESS COUNT
                $filecontent = $word.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".[xls|xlsx]" )
            {
                # OPEN IN EXCEL, PROCESS COUNT
                $filecontent = $excel.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Excel.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".[ppt|pptx]" )
            {
                # OPEN IN POWRPOINT, PROCESS COUNT
                $filecontent = $powerpoint.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Powerpoint.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".pdf" )
            {
                # COUNT WORDS IN PDF FILE
                [int]$wordcount = (Get-Content "$ORIG\$newname" | Measure-Object –Word).Words
            }

            elseif ("$newname" -match ".[txt|csv]" )
            {
                # COUNT WORDS IN TXT FILE
                [int]$wordcount = (Get-Content "$ORIG\$newname" | Measure-Object –Word).Words
            }
            else
            {
                # IDK
                [int]$wordcount = 0
            }
        
            # USE THE WORDCOUNT
            [int]$totalcount += $wordcount
            Write-Output "Wordcount: $wordcount"
            Write-Output "$newname;$wordcount" | Out-File -FilePath "$INFO\$ANALYSIS" -Append 


            #CLOSE FILE
            $filecontent.Close()
        } #>

    } # End of loop processing all source file

} # End of If we have source files



#==========================
#= START A TRADOS PROJECT =


# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.Checked)
{
	Write-Output "Starting Trados Studio..."
    # May not be where expected
    try {
        Set-Location "C:\Program Files (x86)\Trados\Trados Studio\Studio17"
        }
    catch { 
        $TRADOSDIR = (Get-ChildItem -Path "C:\Program Files (x86)" -Filter *.sdlproj -Recurse -ErrorAction SilentlyContinue -Force -File).Directory.FullName
        Write-Output "[DETECTED] Trados in $TRADOSDIR"
        Set-Location $TRADOSDIR
        }

    .\SDLTradosStudio.exe /createProject /name $PROJECTNAME
}


# OK NOW WE WORK
if ($CheckIfOpenExplorer.Checked )
{
    Write-Output "Starting Explorer..."
    start-process explorer "$BASEFOLDER"
}


if ($CheckIfNotify.Checked )
{
    # Have a NICE NOTIFICATION THIS IS BALLERS
    # WOOOOHOOOO
    $objNotifyIcon                      = New-Object System.Windows.Forms.NotifyIcon
    #$objNotifyIcon.Icon = "M:\4_BE\06_General information\Stella\Skrivanek-Rocketlaunch\assets\Rocketlaunch-Icon.ico"  
    $objNotifyIcon.Icon                 = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
    $objNotifyIcon.BalloonTipTitle      = "Fertig!"
    $objNotifyIcon.BalloonTipIcon       = "Info"
#   $objNotifyIcon.BalloonTipText       = -join("Fertig !")
    $objNotifyIcon.Visible              = $True
    $objNotifyIcon.ShowBalloonTip(10000)
}