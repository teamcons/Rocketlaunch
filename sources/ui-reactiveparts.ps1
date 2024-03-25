




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
                            $labelsourcefiles.Text = $text_label_from_Downloads }

        $text_DragNDrop {         
                            $labelsourcefiles.Text = $text_label_DragNDrop }
                            
        $text_nofilesource {
                            $labelsourcefiles.Text = $text_label_nofilesource }
    }
})





#========================================
# Adapting prediction

# Try to guess client from selected email sendermail

#    #[System.Windows.MessageBox]::Show($sourcefiles.SelectedIndex,"Rocketlaunch",1,"Error")



#$sourcefiles.Add_SelectedIndexChanged({
#    [System.Windows.MessageBox]::Show("nuh uh","Nein",1,"Error")
 #   Adapt-Prediction
#})


$sourcefiles.Add_Click({
    Adapt-Prediction
    $gui_code

})








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
#$templates.Rows[1].Selected = $true



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




<# 


 
Add-Type -AssemblyName System.Windows.Forms
    
$UI_Splash                                      = New-Object System.Windows.Forms.Form
$UI_Splash.Text                                 = "Processing data..."
$UI_Splash.Width                                = 310
$UI_Splash.Height                               = 200
$UI_Splash.FormBorderStyle                      = "FixedSingle"
$UI_Splash.ControlBox                           = $false


# FANCY ICON
$UI_Splash_logo                                 = new-object Windows.Forms.PictureBox
$UI_Splash_logo.Width                           = 64
$UI_Splash_logo.Height                          = 64
$UI_Splash_logo.Image                           = $image
$UI_Splash_logo.Location                        = New-Object System.Drawing.Point(100,10)

$progressLabel                                  = New-Object System.Windows.Forms.Label
$progressLabel.Location                         = New-Object System.Drawing.Point(10,84)
$progressLabel.Size                             = New-Object System.Drawing.Size(280, 20)
$progressLabel.Text                             = "0% Complete"
$UI_Splash.Controls.Add($progressLabel)

$progressBar                                    = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location                           = New-Object System.Drawing.Point(10, 104)
$progressBar.Size                               = New-Object System.Drawing.Size(280, 20)
#$progressBar.UseVisualStyleBackColor            = $true
$UI_Splash.Controls.Add($progressBar)

#$UI_Splash.Show()

#$progressBar.Value = 100 ; $progressLabel.Text = "Ready to go !"
#$progressForm.Close()
 #>