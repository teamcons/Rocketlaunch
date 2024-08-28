



        #========================================================
        #                MAIN - Project Creation                =
        #========================================================

<# 
This is where everything happens

This is the main behind the scene thing, which just acts on everything defined before
It grabs from the GUI

#>



#========================================
# Everything. All Of it. All at once.
function Main-ProjectCreation {


        #=================================================
        #                Process Le Input                =
        #=================================================


    #========================================
    # "Close" the form, for the psychological effect of "omg it started"
    # No reaction on the form when starting a new project, is very jarring
    $GUI_Form_MainWindow.WindowState = "Minimized"


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
            
            # If from Outlook, use the dedicated function
            $text_from_Outlook {
                Write-Host "Saving from outlook"
                Save-OutlookAttach $allgoodmails[$sourcefiles.SelectedItems.Index] $ORIG
            }

            # If from downloads, iterate through items, move the checked ones
            $text_from_Downloads {
                Write-Host "Saving from Downloads"
                foreach ( $file in $sourcefiles.Items)
                {
                    if ($file.Checked)
                    {
                        # Path to file is displayed as last item, just use that
                        Write-Output (-join("[MOVE] File at ",$file.SubItems[-1].text))
                        Move-Item -path $file.SubItems[-1].text -Destination $ORIG
                    }

                }
            }

            # If from downloads, iterate through items, move the checked ones
            $text_DragNDrop {
                Write-Host "From DragNDrop"
                foreach ( $file in $sourcefiles.Items)
                {
                    if ($file.Checked)
                    {
                        # Path to file is displayed as last item, just use that
                        Write-Output (-join("[MOVE] File at ",$file.SubItems[-1].text))
                        Move-Item -path $file.SubItems[-1].text -Destination $ORIG
                    }

                }
            }

            # No source, do nothing lol
            $text_nofilesource {
                Write-Host "No source - THIS SHOULD HAVE BEEN FILTERED OUT BY IF"
            }

            # The fuck
            default {
                Write-Host -join ("IDK, WTF IS ",$gui_filesource.SelectedItem)
            }
        } # End of Switch Case



        # Before processing each source file, deal with the archives first
        # Just expand all archives
        Get-ChildItem -Path $ORIG -Filter *.zip -File | Expand-Archive -DestinationPath $ORIG  | Out-Null

        # And get rid of them bc not useful
        Get-ChildItem -Path $ORIG -Filter *.zip -File | Remove-Item        

        # Before processing each source file, move PO and analysis to INFO folder
        Get-ChildItem -Path $ORIG -Filter *Analys*.csv | Move-Item -Destination $INFO
        Get-ChildItem -Path $ORIG -Filter CA-*.pdf | Move-Item -Destination $INFO


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
    if ($gui_folderinoutlook.Text -notmatch $text_nooutlook)        { Create-OutlookFolder $PROJECTNAME $ns $gui_folderinoutlook.Text}

    # Start trados project creator and fill what we can
    if ($CheckIfTrados.Checked)                     { Start-TradosProject $PROJECTNAME $ORIG}

    # Open explorer if its wanted
    if ($CheckIfOpenExplorer.Checked)               { start-process explorer "$BASEFOLDER" }

    # Yeah i redid a Linux command deal with it
    if ($CheckIfNotify.Checked)                     { Notify-Send $PROJECTNAME $text_NotifyText }


    # If user want their homemade changes on the template to be saved
    if ($CheckIfSaveTemplateChanges.Checked)        { Save-DataGridView $templates $templatefile}


    # If user asked to count number of words, and theres actually source files
    if (($CheckIfCountWords.Checked) -and ($gui_filesource.SelectedItem.ToString() -ne $text_nofilesource))        { Count-AllWords $ORIG $INFO}


    # If user want to close app after creation
    if ($CheckIfCloseAfter.Checked)
    {
        Close-All $GUI_Form_MainWindow
    }
    else {

        # If not, recalculate/repredict
        [int]$script:PREDICT_CODE           = (Predict-StructCode)[-1]     
        [void]$gui_code.Items.Add((-join(($PREDICT_CODE),"_")))
        [void]$gui_code.Items.Add((-join(($PREDICT_CODE + 1),"_")))
        [void]$gui_code.Items.Add((-join(($PREDICT_CODE + 2),"_")))
        [void]$gui_code.Items.Add((-join(($PREDICT_CODE + 3),"_")))
        [void]$gui_code.Items.Add((-join(($PREDICT_CODE + 4),"_")))
        $gui_code.SelectedItem              = $gui_code.Items[0]    
    }

}


