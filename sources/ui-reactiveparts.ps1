




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


$gui_filesource.Add_SelectedIndexChanged({
    Write-Output "[UI] CHANGE DETECTED"

    switch ($gui_filesource.SelectedItem ) {
        $text_from_Outlook {
                            $labelsourcefiles.Text = $text_label_from_Outlook }

        $text_from_Downloads {
                            [System.Windows.MessageBox]::Show("Changed","Rocketlaunch",1,"Error")
                            $labelsourcefiles.Text = $text_label_from_Downloads }

        $text_DragNDrop {
                            [System.Windows.MessageBox]::Show("Changed","Rocketlaunch",1,"Error")            
                            $labelsourcefiles.Text = $text_label_DragNDrop }
                            
        $text_nofilesource {
                            $labelsourcefiles.Text = $text_label_nofilesource }
    }
})





#========================================
# Adapting prediction

# Try to guess client from selected email sendermail

#$sourcefiles.Add_Click({Adapt-Prediction})


$sourcefiles_ItemSelectionChanged={
    [System.Windows.MessageBox]::Show("Changed","Rocketlaunch",1,"Error")
    #Adapt-Prediction
}
$sourcefiles.Add_Mouseclick({echo "MOUSE"})
$sourcefiles.Add_MouseUp({echo "MOUSE UP"})
$sourcefiles.Add_Click($sourcefiles_ItemSelectionChanged)
$sourcefiles.Add_SelectedIndexChanged({echo "CHANGED"})









#========================================
# Load templates



# Fill the grid




for ($i=1; $i -lt $templates.ColumnCount ; $i++)
{
    if ($i -le 10)  {$templates.Columns[$i].Name = -join("0",($i-1)) }
    else            {$templates.Columns[$i].Name = ($i-1)}

    $templates.Columns[$i].Width = 80
}



# Try to load templates, and if that doesnt work, have minimal ones
$templates = load_template $templates $templatefile


# Select the second one in the list 
# The first one is way too skeleton
$templates.Rows[1].Selected = $true #.Selected = $true




#========================================
# Button clicks

# If checked, topmost is checked (and window then stays on top)
$gui_keepontop.Add_Click({$GUI_Form_MainWindow.Topmost = $gui_keepontop.Checked})

# Reload the relevant emails
$sourcefile_refreshButton.Add_Click({$sourcefiles.Items.Clear() ; Load-RelevantMails})

# Show settings
$gui_help.add_click({$GUI_Form_MoreStuff.ShowDialog()})

# Launch project creation
$gui_okButton.add_click({Main-ProjectCreation})

# Close everything
$gui_cancelButton.add_click({Close-All $GUI_Form_MainWindow})


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