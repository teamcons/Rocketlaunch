

# Look for emails with attachments
# For each email, we look attachments and count the ones with supported formats
# We are not interested in junk like image001.jpg etc... which is signatures and stuff

Write-Output "[START] Loading Outlook capabilities"


#================================================================
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


#================================================================
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

} # End of function

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



#================================================================
# If yesterday was sunday, it was, in fact, friday all along.
function Get-LastBusinessDay
{
    #Day of the week
    # Sonntag = 0, Samstag = 6
    # If we are monday, grab emails starting last friday
    
    if ( (Get-Date -UFormat "%u") -eq 1 )
    {
        $date = (Get-Date (Get-Date).AddDays(-3) -Format 'dd/MM/yyyy 17:30')
    }
    else {
        $date = (Get-Date (Get-Date).AddDays(-1) -Format 'dd/MM/yyyy 17:30')
    }
    
    return $date    
}
    




#================================================================
#### GET RELEVANT MAILSSSS
function Load-RelevantMails
{


    [bool]$StopSearching = $false
    $script:allgoodmails = New-Object -TypeName 'System.Collections.ArrayList'

    foreach ($mail in $allmails)
    {
    
        [bool]$AddToGoodMails = $false
        [int]$CountGoodAttachments = 0
        foreach ( $attach in $mail.Attachments ) 
        {
            #echo $attach.FileName
            if ($attach.FileName -match  ".(pdf|doc|docx|xls|xlsx|ppt|pptx|xml|idml|csv|txt|zip|sdlppx)" )
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
            [void]$sourcefilesItem.Subitems.Add( (Get-Date -Date $mail.ReceivedTime -UFormat "%A, %H:%M").ToString() )
            [void]$sourcefiles.Items.Add($sourcefilesItem)
            $goodmailindex += 1
        } # End of adding goodmail
    
    } # End of looking for emails with attachments
    
    

} # End of function Load-RelevantMails
    



#================================================================
# If theres something, define a default selected

function Adapt-Prediction {

    [int]$selected                      = $sourcefiles.SelectedItems[0].Index
    [string]$email                      = $allgoodmails[$selected].SenderEmailAddress
    [string]$FirstSelectionMail         = (Get-CompanyName $email)[-1]

    # Conserve the code user was using and if they entered any additional text after client name
    #$afterclientname = ($gui_code.Text.split(' ',2))[1]
    #if ($afterclientname -ne $none) {$afterclientname = -join(" ",$afterclientname)}

    [string]$gui_code.Items[0] = -join($PREDICT_CODE,"_",$FirstSelectionMail)
    [string]$gui_code.Items[1] = -join(($PREDICT_CODE + 1),"_",$FirstSelectionMail)
    [string]$gui_code.Items[2] = -join(($PREDICT_CODE + 2),"_",$FirstSelectionMail)
    [string]$gui_code.Items[3] = -join(($PREDICT_CODE + 3),"_",$FirstSelectionMail)
    [string]$gui_code.Items[4] = -join(($PREDICT_CODE + 4),"_",$FirstSelectionMail)

    $gui_code.Text = -join($gui_code.Text.split("_",2)[0],"_",$FirstSelectionMail)
}

#================================================================

$OL                         = New-Object -ComObject OUTLOOK.APPLICATION
$ns                         = $OL.GETNAMESPACE("MAPI")
$date                       = Get-LastBusinessDay
$filter                     = "[ReceivedTime] >= '$date'"
#$filter                     = "[ReceivedTime] >= '$date' And [FlagStatus] = 6"
#$filter                     = query ="@SQL='urn:schemas:httpmail:hasattachment'=1"
$allmails                   = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)


Load-RelevantMails
try     {$sourcefiles.Items[0].Selected = $true ; Adapt-Prediction}
catch   {Write-Output "No mail with relevant attach !"}




