

# Look for emails with attachments
# For each email, we look attachments and count the ones with supported formats
# We are not interested in junk like image001.jpg etc... which is signatures and stuff

Write-Output "[START] Loading Outlook capabilities"

# Folder in outlook
function Create-OutlookFolder {
    param(
        [string]$foldername,
        $namespace
    )
    Write-Output "[CREATE] Folder in Outlook"
    [string]$Username = $Env:UserName.split(".")[0]
    $TextInfo = (Get-Culture).TextInfo
    [string]$Username = $TextInfo.ToTitleCase($Username)
    $namespace.Folders.Item(1).Folders.Item("Posteingang").Folders.Item("02_ONGOING JOBS").Folders.Item($Username).Folders.Add($PROJECTNAME)
}

# Takes a mail, takes a destination, put every relevant attach there (ignore signatures and similar BS)
function Save-OutlookAttach {
    param($mail,$destination)

    foreach ($attachment in $mail.Attachments)
    {
        if ($attachment.FileName -notmatch "^image[0-9][0-9][0-9]")
        {
            Write-Output (-join($destination,"\",$attachment.FileName))
            $attachment.SaveAsFile( -join($destination,"\",$attachment.FileName) )

        }
    } # End of attachment processing

}

<#     elseif ($gui_filesource.SelectedItem -match $text_from_Downloads )
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
    } # End of user load themselves #>






$OL                         = New-Object -ComObject OUTLOOK.APPLICATION
$ns                         = $OL.GETNAMESPACE("MAPI")
$date                       = Get-Date (Get-Date).AddDays(-1) -Format 'dd/MM/yyyy HH:mm'
$filter                     = "[ReceivedTime] >= '$date'"
#$filter                     = "[ReceivedTime] >= '$date' And [FlagStatus] = 6"
#$filter                     = query ="@SQL='urn:schemas:httpmail:hasattachment'=1"
$allmails                   = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)




#### GET RELEVANT MAILSSSS

[bool]$StopSearching = $false
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
    
    


    

# If theres something, define a default selected
try { 
    $sourcefiles.Items[0].Selected = $true 
}
catch {
    Write-Output "No mail with relevant attach !"
}

function Add-Info-To-Combobox{
    param($combobox)
    echo no
}

[int]$PREDICT_CODE = (Predict-StructCode)[-1] 
[String]$FirstSelectionMail = (Get-CompanyName $allgoodmails.Item(0).SenderEmailAddress)[-1]

[string]$gui_code.Items.Add(-join($PREDICT_CODE,"_",$FirstSelectionMail ))
[string]$gui_code.Items.Add(-join(($PREDICT_CODE + 1),"_",$FirstSelectionMail ))  
[string]$gui_code.Items.Add(-join(($PREDICT_CODE + 2),"_",$FirstSelectionMail ) ) 
[string]$gui_code.Items.Add(-join(($PREDICT_CODE + 3),"_",$FirstSelectionMail ))

$gui_code.SelectedItem = $gui_code.Items[0]