


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
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

# Load assets
$script:icon                = New-Object system.drawing.icon $ScriptPath\assets\icon.ico
$script:templatefile        = -join($ScriptPath,"\documentation\Project templates.csv")
$script:image               = [system.drawing.image]::FromFile((get-item $ScriptPath\assets\icon-mini.ico))

# Load everything we need
Import-Module $ScriptPath/sources/text.ps1
Import-Module $ScriptPath/sources/defaults.ps1
Import-Module $ScriptPath/sources/internals.ps1
Import-Module $ScriptPath/sources/ui-MainWindow.ps1 
Import-Module $ScriptPath/sources/ui-SettingsDialog.ps1 
Import-Module $ScriptPath/sources/outlook-backend.ps1 




#========================================
# Create the Emails table
$Datatable_Emails = New-Object System.Data.DataTable
$newcol = New-Object system.Data.DataColumn $text_columns_Subject,([string]); $Datatable_Emails.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_Sendername,([string]); $Datatable_Emails.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_Attachments,([int]); $Datatable_Emails.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_time,([int]); $Datatable_Emails.columns.add($newcol)  

# Create the Files In Downloads table
$Datatable_FilesInDownloads = New-Object System.Data.DataTable
$newcol = New-Object system.Data.DataColumn "Checked",([bool]); $Datatable_FilesInDownloads.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_DL_File,([string]); $Datatable_FilesInDownloads.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_DL_LastWrite,([string]); $Datatable_FilesInDownloads.columns.add($newcol)  


# Create the Files DragNDrop table
$Datatable_FilesDragNDrop = New-Object System.Data.DataTable
$newcol = New-Object system.Data.DataColumn "Checked",([bool]); $Datatable_FilesDragNDrop.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_DL_File,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_DD_Path,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)  





        #=======================================================
        #                Display User Interface                =
        #=======================================================



#========================================
# Interface defined in the ui module
Write-Output "[START] Show main window"; $result = $GUI_Form_MainWindow.ShowDialog()
#$result = $GUI_Form_MainWindow.Show()
#$GUI_Form_MainWindow.Activate()


# Cancel culture : Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) { Write-Output "[INPUT] Got Cancel. Aw. Exit." ; exit }



        #=================================================
        #                Process Le Input                =
        #=================================================


#========================================
# Make sure we have clean input
[string]$PROJECTNAME                = $gui_code.Text ; Write-Output "[INPUT] Got: $PROJECTNAME"
[string]$PROJECTNAME                = (Get-CleanifiedCodename $PROJECTNAME)[-1]
[string]$BASEFOLDER                 = (Rebuild-Tree $PROJECTNAME)[-1]

# Create project folder
Write-Output "[ACTION] Create base folder: $BASEFOLDER"
New-Item -ItemType Directory -Path "$BASEFOLDER"

# Get selected element. Skip the first element cuz no
$selectedrow                        = $templates.CurrentCell.RowIndex
$allfolderstocreate                 = ($templates.Rows[$selectedrow].Cells | Select-Object -Skip 1 )

# CREATE ALLLLL THE FOLDERS
Create-AllFolders $BASEFOLDER $allfolderstocreate





        #=======================================================
        #                Include Original Files                =
        #=======================================================


#========================================
# If user asked to include source files, include those in new folder, with naming conventions
if ($gui_filesource.SelectedItem.ToString() -ne $text_nofilesource)
{


    # CHECK WE HAVE THE MINIMUM FOLDERS BECAUSE WE DONT KNOW WHAT TEMPLATE USER USED
    # IF THE STANDARD MINIMUM ISNT THERE, JUST USE BASE FOLDER INSTEAD
    [string]$INFO = -join($BASEFOLDER,"\",(Get-ChildItem -Path "$BASEFOLDER" -Filter "00_*" | Select-Object -First 1).Name)
    [string]$ORIG = -join($BASEFOLDER,"\",(Get-ChildItem -Path "$BASEFOLDER" -Filter "01_*" | Select-Object -First 1).Name)


    # Check which text has the combobox to decide how to handle this.
    switch ($gui_filesource.SelectedItem) {
        $text_from_Outlook {
            Write-Host "Saving from outlook"
            Save-OutlookAttach $allgoodmails[$sourcefiles.SelectedItems.Index] $ORIG
        }
        $text_from_Downloads {
            Write-Host "From Downloads, not implemented yet !"
        }
        $text_DragNDrop {
            Write-Host "From DragNDrop, not implemented yet !"
        }
        $text_nofilesource {
            Write-Host "No source - THIS SHOULD HAVE BEEN FILTERED OUT BY IF"
        }
        default {
            Write-Host -join ("IDK, WTF IS ",$gui_filesource.SelectedItem)
        }
    } # End of Switch Case

    # Before processing each source file, deal with the archives first
    # Just expand all archives
    Get-ChildItem -Path $ORIG -Filter *.zip | Expand-Archive -DestinationPath $ORIG
   
    # Make sure everything saved is named as we need it
    # Convention is to have Projectcode-File_orig.fileext
    Rename-Source $ORIG $PROJECTNAME.Substring(0,9) "_orig"

} # End of If we have source files





        #===============================================
        #                POSTPROCESSING                =
        #===============================================


#========================================
# Pin to quick access in explorer
if ($CheckIfCreateExplorerQuickAccess.Checked)  { Create-QuickAccess $BASEFOLDER }

# Create a folder in outlook
if ($CheckIfCreateOutlookFolder.Checked)        { Create-OutlookFolder $PROJECTNAME $ns }

# Start trados project creator and fill what we can
if ($CheckIfTrados.Checked)                     { Start-TradosProject $PROJECTNAME }

# Open explorer if its wanted
if ($CheckIfOpenExplorer.Checked)               { start-process explorer "$BASEFOLDER" }

# Yeah i redid a Linux command deal with it
if ($CheckIfNotify.Checked)                     { Notify-Send $PROJECTNAME $text_NotifyText }


# Create an application context for it to all run within. 
# This helps with responsiveness and threading.
#$appContext = New-Object System.Windows.Forms.ApplicationContext 
#[void][System.Windows.Forms.Application]::Run($appContext)