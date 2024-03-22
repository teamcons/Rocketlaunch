
# Everything. All Of it. All at once.
function Main-ProjectCreation {


#Import-Module $ScriptPath\modules\PoshTaskbarItem
#$ti = New-TaskbarItem -Title 'Countdown'
#Show-TaskbarItem $ti

#Set-TaskbarItemProgressIndicator $ti -Progress 1 -State Paused

# "Close" the form, for the psychological effect of "omg it started"
# No reaction on the form when starting a new project, is very jarring
$GUI_Form_MainWindow.Hide()



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
            # Foreach path in $sourcefiles.SelectedItems.Value
            # Move-Item -path $path -Destination $ORIG
        }
        $text_DragNDrop {
            Write-Host "From DragNDrop, not implemented yet !"
            # Foreach path in $sourcefiles.SelectedItems.Value
            # Move-Item -path $path -Destination $ORIG

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
    Get-ChildItem -Path $ORIG -Filter *.zip | Expand-Archive -DestinationPath $ORIG  | Out-Null

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


# If user want their homemade changes on the template to be saved
if ($CheckIfSaveTemplateChanges.Checked)        { Save-DataGridView $templates $templatefile}


# If user want their homemade changes on the template to be saved
if ($CheckIfCountWords.Checked)        { Count-AllWords $ORIG $INFO}


# If user want to close app after creation
if ($CheckIfCloseAfter.Checked)                 { Close-All $GUI_Form_MainWindow}
else {$GUI_Form_MainWindow.Show()}


}


