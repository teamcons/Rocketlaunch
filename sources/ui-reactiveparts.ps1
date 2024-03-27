




#========================================
# Code prediction

[int]$script:PREDICT_CODE           = (Predict-StructCode)[-1]     
[void]$gui_code.Items.Add((-join(($PREDICT_CODE),"_")))
[void]$gui_code.Items.Add((-join(($PREDICT_CODE + 1),"_")))
[void]$gui_code.Items.Add((-join(($PREDICT_CODE + 2),"_")))
[void]$gui_code.Items.Add((-join(($PREDICT_CODE + 3),"_")))
[void]$gui_code.Items.Add((-join(($PREDICT_CODE + 4),"_")))
$gui_code.SelectedItem              = $gui_code.Items[0]



#========================================
# Change label if file source has changed
# rebuild everything depending on what selection has been

# Create the Files DragNDrop table
$Datatable_FilesDragNDrop = New-Object System.Data.DataTable
#$newcol = New-Object system.Data.DataColumn "Checked",([bool]); $Datatable_FilesDragNDrop.columns.add($newcol) 
$newcol = New-Object system.Data.DataColumn $text_columns_File,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)  
$newcol = New-Object system.Data.DataColumn $text_columns_Directory,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn $text_columns_LastWrite,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn $text_columns_Path,([string]); $Datatable_FilesDragNDrop.columns.add($newcol)

#========================================
function Rebuild-Outlook-View
{
    [void]$sourcefiles.Columns.Add($text_columns_Subject,330)
    [void]$sourcefiles.Columns.Add($text_columns_Sendername,200)
    [void]$sourcefiles.Columns.Add($text_columns_Attachments,80)
    [void]$sourcefiles.Columns.Add($text_columns_time,100)
    Load-RelevantMails
} # End of Rebuild-Outlook-View


#========================================
function Rebuild-DragNDrop-View
{
    #[void]$sourcefiles.Columns.Add("Checked",100)
    [void]$sourcefiles.Columns.Add($text_columns_File,180)
    [void]$sourcefiles.Columns.Add($text_columns_Directory,100)
    [void]$sourcefiles.Columns.Add($text_columns_LastWrite,160)
    [void]$sourcefiles.Columns.Add($text_columns_Path,260)

    foreach ($row in $Datatable_FilesDragNDrop.rows)
    {
        $sourcefilesItem            = New-Object System.Windows.Forms.ListViewItem($row[$text_columns_File])
        $sourcefilesItem.Checked    = $true
        [void]$sourcefilesItem.Subitems.Add($row[$text_columns_Directory])
        [void]$sourcefilesItem.Subitems.Add($row[$text_columns_LastWrite])
        [void]$sourcefilesItem.Subitems.Add($row[$text_columns_Path])
        [void]$sourcefiles.Items.Add($sourcefilesItem)
    }
} # End of Rebuild-DragNDrop-View

#========================================
function Rebuild-Downloads-View
{
    #[void]$sourcefiles.Columns.Add("Checked",100)
    [void]$sourcefiles.Columns.Add($text_columns_File,180)
    [void]$sourcefiles.Columns.Add($text_columns_Directory,100)
    [void]$sourcefiles.Columns.Add($text_columns_LastWrite,160)
    [void]$sourcefiles.Columns.Add($text_columns_Path,260)

    # For each file in the downloads folder
    foreach ($file in (Get-ChildItem $env:USERPROFILE\Downloads  | Sort LastWriteTime -Descending ))
    {
        # If its fresh from today
        if ($file.LastWriteTime.ToString("dd.MM") -match (Get-Date -Format "dd.MM"))
        {
            # Add it
            $sourcefilesItem            = New-Object System.Windows.Forms.ListViewItem($file.Name)
            $sourcefilesItem.Checked    = $false
            [void]$sourcefilesItem.Subitems.Add($file.Directory.Name)
            [void]$sourcefilesItem.Subitems.Add($file.LastWriteTime.ToString())
            [void]$sourcefilesItem.Subitems.Add($file.FullName)
            [void]$sourcefiles.Items.Add($sourcefilesItem)
        }
    }
} # End of Rebuild-Downloads-View





#========================================
Function Reset-View{

    # Empty everything
    Write-Output "[UI] CHANGE DETECTED"
    $sourcefiles.Items.Clear()
    $sourcefiles.Columns.Clear()

    # Rebuild based on what is the source
    switch ($gui_filesource.Text ) {

        # User selected Outlook
        $text_from_Outlook {
                            $labelsourcefiles.Text      = $text_label_from_Outlook
                            $sourcefiles.Checkboxes     = $false
                            Rebuild-Outlook-View
                        }
        # User selected Downloads
        $text_from_Downloads {
                            $labelsourcefiles.Text      = $text_label_from_Downloads
                            $sourcefiles.Checkboxes     = $true 
                            Rebuild-Downloads-View
                        }
        # User selected DragNDrop
        $text_DragNDrop {
                            $labelsourcefiles.Text      = $text_label_DragNDrop
                            $sourcefiles.Checkboxes     = $true 
                            Rebuild-DragNDrop-View
                        }
        # User selected no
        $text_nofilesource {
                            $labelsourcefiles.Text      = $text_label_nofilesource
                            $sourcefiles.Checkboxes     = $false                
                        }
    }# End of Switch
}





#================================
# When user needs to reload the view
$sourcefile_refreshButton.Add_Click({Reset-View})

#================================
# When user select a different source
$gui_filesource.Add_SelectedIndexChanged({Reset-View})










#================================
# Signal that dropping files work
$DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}


#================================
# When a file is dragged onto sourcefiles, add it to it
$DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filepath in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
    {
        $file = Get-Item $filepath
        $file
        #$ico =  ([System.Drawing.Icon]::ExtractAssociatedIcon($filepath) ).ToBitmap()
        #$sourcefiles.Rows.Add($true,$file.Name,$file.LastWriteTime,$file.Directory.FullName);

        $row = $Datatable_FilesDragNDrop.NewRow()
        #$row[0] = $true
        $row[$text_columns_File] = $file.Name
        $row[$text_columns_Directory] = $file.Directory.Name
        $row[$text_columns_LastWrite] = $file.LastWriteTime.ToString()
        $row[$text_columns_Path] = $file.FullName
        $Datatable_FilesDragNDrop.rows.Add($row)

        # Correct view, add new item
        if ($gui_filesource.Text -match $text_DragNDrop)
            {$sourcefilesItem = New-Object System.Windows.Forms.ListViewItem($file.Name)
            $sourcefilesItem.Checked = $true
            [void]$sourcefilesItem.Subitems.Add($file.Directory.Name)
            [void]$sourcefilesItem.Subitems.Add($file.LastWriteTime.ToString())
            [void]$sourcefilesItem.Subitems.Add($file.FullName)
            [void]$sourcefiles.Items.Add($sourcefilesItem)
        }



	} # End of processing list
}

$sourcefiles.Add_DragOver($DragOver)
$sourcefiles.Add_DragDrop($DragDrop)









<# 
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
 #>



#========================================
# Adapting prediction
# Try to guess client from selected email sendermail
$sourcefiles.Add_Click({ 
    if ($gui_filesource -match $text_from_Outlook)
        {Adapt-Prediction}
})



#========================================
# Load templates

# Try to load templates, and if that doesnt work, have minimal ones
$templates = load_template $templates $templatefile


# Select the second one in the list 
# The first one is way too skeleton
#$templates.Rows[1].Selected = $true



#========================================
# Button clicks

# If checked, topmost is checked (and window then stays on top)
$gui_keepontop.Add_Click({$GUI_Form_MainWindow.Topmost = $gui_keepontop.Checked})



# Show settings
$gui_help.add_click({$GUI_Form_MoreStuff.ShowDialog()})

# Launch project creation
$gui_okButton.add_click({Main-ProjectCreation})

# Close everything
$gui_cancelButton.add_click({Close-All $GUI_Form_MainWindow})

