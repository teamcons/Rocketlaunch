#================================================================================================================================



        #===============================================
        #                Initialization                =
        #===============================================


#========================================
# Fancy !
Write-Output "  ================================"
Write-Output "  =        -ROCKETLAUNCH!        ="
Write-Output "  ================================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Start new projects, but very very quickly !"
Write-Output "GPL-3.0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output ""
Write-Output ""


#========================================
# Grab script location in a way that is compatible with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }




#========================================
# Get all resources

Import-Module $ScriptPath/text.ps1
Import-Module $ScriptPath/defaults.ps1
Import-Module $ScriptPath/bigstring.ps1
Import-Module $ScriptPath/internals.ps1
Import-Module $ScriptPath/ui-MainWindow.ps1 
Import-Module $ScriptPath/ui-SettingsDialog.ps1 
Import-Module $ScriptPath/outlook-backend.ps1 




        #=======================================================
        #                Display User Interface                =
        #=======================================================



#========================================
# Interface defined in the ui module
Write-Output "[START] Show main window"
$result = $GUI_Form_MainWindow.ShowDialog()

# Cancel culture : Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    { Write-Output "[INPUT] Got Cancel. Aw. Exit." ; exit }


# penis
[string]$PROJECTNAME        = $gui_code.Text 
Write-Output "[INPUT] Got: $PROJECTNAME"




        #=================================================
        #                Process Le Input                =
        #=================================================


#========================================
# Make sure we have clean input
[string]$PROJECTNAME                = (Get-CleanifiedCodename $PROJECTNAME)[-1]
[string]$BASEFOLDER                 = (Rebuild-Tree $PROJECTNAME)[-1]

# Create project folder
Write-Output "[ACTION] Create base folder: $BASEFOLDER"
New-Item -ItemType Directory -Path "$BASEFOLDER"

# CREATE ALLLLL THE FOLDERS
# Get selected element. Skip the first element cuz no
$selectedrow                        = $templates.CurrentCell.RowIndex
$allfolderstocreate                 = ($templates.Rows[$selectedrow].Cells | Select-Object -Skip 1 )
Create-AllFolders $BASEFOLDER $allfolderstocreate

# PIN TO EXPLORER
if ($CheckIfCreateExplorerQuickAccess.Checked) { Create-QuickAccess $BASEFOLDER }

# Outlook folder
if ($CheckIfCreateOutlookFolder.Checked) { Create-OutlookFolder $PROJECTNAME $ns }




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


    #if ($gui_filesource.SelectedItem -match $text_from_Outlook )
    #{#} # End of process outlook inclusion

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
    }



    # Before processing each source file, deal with the archives first
    if ($CheckIfExpandArchives.Checked) { Get-ChildItem -Path $ORIG -Filter *.zip | Expand-Archive -DestinationPath $ORIG }
   

    # Make sure everything saved is named as we need it
    Rename-Source $ORIG $PROJECTNAME.Substring(0,9) "_orig"



} # End of If we have source files


#========================================
# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.Checked) { Start-TradosProject $PROJECTNAME }

# OK NOW WE WORK
if ($CheckIfOpenExplorer.Checked) { start-process explorer "$BASEFOLDER" }

# Yeah i redid a Linux command deal with it
if ($CheckIfNotify.Checked) { Notify-Send $PROJECTNAME $text_NotifyText }


# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
#$appContext = New-Object System.Windows.Forms.ApplicationContext
#[void][System.Windows.Forms.Application]::Run($appContext)