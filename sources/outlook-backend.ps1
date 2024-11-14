




        #===============================================
        #                OUTLOOK BACKEND               =
        #===============================================

<# 
Everything Outlook and to load emails
Contains stuff executed on load
#>



#================================================================
# Look for emails with attachments
# For each email, we look attachments and count the ones with supported formats
# We are not interested in junk like image001.jpg etc... which is signatures and stuff

Write-Output "[START] Loading Outlook capabilities"


#================================================================
# Folder in outlook
function Create-OutlookFolder {
    param(
        [string]$foldername,
        $namespace,
        [string]$infolder
    )
    Write-Output "[CREATE] Folder in Outlook"
    [string]$Username = $Env:UserName.split(".")[0]
    $TextInfo = (Get-Culture).TextInfo
    [string]$Username = $TextInfo.ToTitleCase($Username)
    $namespace.Folders.Item(1).Folders.Item("Posteingang").Folders.Item($infolder).Folders.Item($Username).Folders.Add($PROJECTNAME)
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

    $ProgressLabel.Text                 = $text.Splash.loadingoutlook
    $ProgressBar.Value                  = 0

    $OL                                 = New-Object -ComObject OUTLOOK.APPLICATION
    $script:ns                          = $OL.GETNAMESPACE("MAPI")
    $date                               = Get-LastBusinessDay
    $filter                             = "[ReceivedTime] >= '$date'"
    #$filter                            = "[ReceivedTime] >= '$date' And [FlagStatus] = 6"
    #$filter                            = query ="@SQL='urn:schemas:httpmail:hasattachment'=1"
    $script:allmails                    = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)
    
    # So we know how to iterate the splash
    # Ok, outlook itself can count as one email loaded

    [int]$percent_per_email             = (100 / ($allmails.Count() + 1 ) )

    # So we loaded Outlook, count it as progress
    [string]$ProgressBar.Value          = $percent_per_email
    Write-Output (-join("[LOAD] Update Splash: ",$ProgressBar.Value,"% done"))    

    # This holds all the mails relevant to us
    # It avoids us the hassle later
    $script:allgoodmails = New-Object -TypeName 'System.Collections.ArrayList'

    # For each email we get our hands on
    foreach ($mail in $allmails)
    {


        # Saying what we parse on the splash
        # We dont display the full mail subject, but limit to 40 characters
        # To avoid some ugly text clipping below the progress bar
        # If you cant say it in 40 characters, dont say it
        $ProgressLabel.Text = -join($text.Splash.loadingmail,-join($mail.Subject[0..35]),"...")
        Write-Output (-join("[LOAD] Processing: ",$mail.Subject,"..."))

        # Deal with it only if SMTP
        # Answers from us are "EX", and parsing them just is waste of timne
        if ($mail.SenderEmailType -match "SMTP")
        {
        
            # First loop through the email, to see if it has the attachments
            [bool]$AddToGoodMails = $false
            [int]$CountGoodAttachments = 0
            foreach ( $attach in $mail.Attachments ) 
            {
                #echo $attach.FileName
                # Ignore unsupported formats
                if ($attach.FileName -notmatch $script:unsupported )
                #if ($attach.FileName -match $accepted_attachments )
                {
                    echo (-join("MATCH:",$attach.FileName))
                    $AddToGoodMails = $true
                    [int]$CountGoodAttachments += 1
        
                }   # End of each mails
                #else {echo (-join("NOTMATCH:",$attach.FileName)) }  
            } #End of checking attachments
        
            # we found one with a supported attachment !
            # So we add it where we need it. Save it in a list also
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

        } # End of only deal in SMTP



        # Update splash progress bar
        # Avoid going over 100 because we are dealing in Int
        if (($ProgressBar.Value + $percent_per_email) -gt 100)
            {$ProgressBar.Value = 100}
        else
            {$ProgressBar.Value = $ProgressBar.Value + $percent_per_email}

        # We update even if the mail was skipped - Else the progress bar never goes fully until the end
        # Even if it displays half a second, it looks more reactive too
        Write-Output (-join("[LOAD] Update Splash: ",$ProgressBar.Value,"% done"))



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

# Check if theres a "DELIVERED" folder in outlook for the code
# If not create it

function Create-ArchiveFolder {
    param(
        [string]$projectcode,
        $namespace
    )

    # Calculate range
    # "2024-2033_client to 2024-2033"
    $code = $projectcode.Split("_")[0]

    # "2024-2033" to "2024-20"
    $range = $code -replace ".{2}$" 

    # "2024-20" to "2024-2000 to 2024-2099"
    $range = $range + "00 to " + $range + "99"

    $yearfolder = "XTRF-" + $code.Split("-")[0]
    
    try {
        $namespace.Folders.Item(1).Folders.Item("Posteingang").Folders.Item("03_DELIVERED JOBS").Folders.Item($yearfolder).Folders.Item($range)
    }
    catch {
        echo "[CREATE] Archive top folder"
        $namespace.Folders.Item(1).Folders.Item("Posteingang").Folders.Item("03_DELIVERED JOBS").Folders.Item($yearfolder).Folders.Add($range)
    }



}







#================================================================

# Load relevant emails.
Load-RelevantMails

# Try selecting the first item
try     {
    $sourcefiles.Items[0].Selected = $true ; Adapt-Prediction}


# If it does not work, then there is no relevant email at all
catch   {
    Write-Output "No mail with relevant attach !"

    # If there is no email at all, then user may want to switch to loading from downloads
    # So do it for them
    $gui_filesource.Text = $text_from_Downloads
    Reset-View

    # After loading downloads, if there is no item, then user may want to drag and drop instead
    if ($sourcefiles.Items.Count -eq 0) 
    {
        # So switch to exactly that.
        $gui_filesource.Text = $text_DragNDrop
        Reset-View
    }
}


