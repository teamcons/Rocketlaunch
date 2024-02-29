﻿#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1"
# Usage: Compiled form, just double-click.
#
#
#----------------STEPS----------------
#
# Initialization
# GUI
# Processing Input
# Build the project
# Bonus
#
#-------------------------------------


#===============================================
#                Initialization                =
#===============================================


#========================================
# Fancy !
Write-Output "================================"
Write-Output "=        -ROCKETLAUNCH!        ="
Write-Output "================================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Start new projects really, really quick!"
Write-Output "CC0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output ""
Write-Output ""


#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$APPNAME            = "-Rocketlaunch!"

# Load templates from a csv in same place as executable
#[string]$LOAD_TEMPLATES_FROM = $MyInvocation.MyCommand.Path
[string]$ROOTSTRUCTURE      = "M:\9_JOBS_XTRF\"
[regex]$CODEPATTERN         = -join($YEAR,"-[0-9]")
[string]$YEAR               = get-date –f yyyy



# Grab script location in a way that is compatible with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
{ 
    $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
}
else
{ 
    $global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } 
}




#========================================
# Get all resources
Import-Module $ScriptPath/text.ps1
Import-Module $ScriptPath/defaults.ps1
Import-Module $ScriptPath/internals.ps1
Import-Module $ScriptPath/ui.ps1 



#==============================================================
#                GUI - Ask the Right Questions                =
#==============================================================

# Look for emails with attachments
# For each email, we look attachments and count the ones with supported formats
# We are not interested in junk like image001.jpg etc... which is signatures and stuff

    $allgoodmails = New-Object -TypeName 'System.Collections.ArrayList'
    foreach ($mail in $allmails)
    {
        echo $mail.SenderName
    
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
    
    
    #}
    
    
    # Add the ListView to the Form
    try { 
        $sourcefiles.Items[0].Selected = $true 
    }
    catch {
        Write-Output "No mail with relevant attach !"
    }
    
    try { 
        $allgoodmails.Item(0).SenderEmailAddress -match "@(?<content>.*).com"
        
        #$attempt_at_companyname         = $matches["content"]
        $attempt_at_companyname         = [cultureinfo]::GetCultureInfo("de-DE").TextInfo.ToTitleCase($attempt_at_companyname)
        echo $attempt_at_companyname
        [string]$gui_code.Items[0] = -join($PREDICT_CODE,"_",$attempt_at_companyname )
        [string]$gui_code.Items[1] = -join(($PREDICT_CODE + 1),"_",$attempt_at_companyname )  
        [string]$gui_code.Items[2] = -join(($PREDICT_CODE + 2),"_",$attempt_at_companyname )  
        [string]$gui_code.Items[3] = -join(($PREDICT_CODE + 3),"_",$attempt_at_companyname )
    }
    catch {
        Write-Output "Messy email !"
    }
    






#==============================================================
#                     Processing Le input                     =
#==============================================================


$result = $GUI_Form_MainWindow.ShowDialog()

[string]$PROJECTNAME        = $gui_code.Text 
$PROJECTTEMPLATE            = $templates.SelectedItems.Text

# Cancel culture
# Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Output "[INPUT] Got Cancel. Aw. Exit."
        exit
    }
Write-Output "[INPUT] Got: $PROJECTNAME"



# Empty, so go on with what was initially predicted
if ("$PROJECTNAME" -notmatch "^[0-9]" )
{

    $PROJECTNAME = -join($PREDICT_CODE,$PROJECTNAME)
    Write-Output "Its words. Now: $PROJECTNAME"
}

# Remove invalid character, just in case
$PROJECTNAME = $PROJECTNAME.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
Write-Output "Removed invalid. Now: $PROJECTNAME"


# is it missing zeros
if ($PROJECTNAME -match "^[0-9][0-9][0-9]")
{
    $PROJECTNAME = -join("0",$PROJECTNAME)
    Write-Output "Missing first zero. Now: $PROJECTNAME"

}
elseif ($PROJECTNAME -match "^[0-9][0-9]")
{
    $PROJECTNAME = -join("00",$PROJECTNAME)
    Write-Output "Missing two zero. Now: $PROJECTNAME"
}
elseif ($PROJECTNAME -match "^[0-9]")
{
    $PROJECTNAME = -join("000",$PROJECTNAME)
    Write-Output "Missing three zero. Now: $PROJECTNAME"
}

# Add year
$PROJECTNAME = -join($YEAR,"-",$PROJECTNAME)




##### Ultimate check
try { $DIRCODE = $PROJECTNAME.SubString(0, 9) }
catch {
	$ERRORTEXT="Projektcode ist unpassend !!!
Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
Angegeben: $PROJECTCODE"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)

exit }



#==============================================================
#                      Build The Project                      =
#==============================================================
# REBUILT THE WHOLE TREE
# Its in... year, underscore
$BASEFOLDER     = -join($ROOTSTRUCTURE,$DIRCODE.Substring(0,4),"_")
$BASEFOLDER     = -join($BASEFOLDER,$DIRCODE.Substring(5,2),"00-",$DIRCODE.Substring(5,2),"99")

# If the folder with project numbers in range do not exist, just create it lol
if (!(Test-Path $BASEFOLDER -PathType Container)) {
    Write-Output "[CREATE] Range folder in tree: $BASEFOLDER"
    New-Item -ItemType Directory -Force -Path "$BASEFOLDER"
}
$BASEFOLDER = -join($BASEFOLDER,"\",$PROJECTNAME)



Write-Output "[CREATE] Base folder: $BASEFOLDER"
New-Item -ItemType Directory -Path "$BASEFOLDER"

# Count folder number
[int]$foldernumber = 0 

# CREATE ALLLLL THE FOLDERS
# Skip the first element cuz no
foreach ($folder in ($templates.Rows[$templates.CurrentCell.RowIndex].Cells | Select-Object -Skip 1 ) )
{
    echo $folder
    if ($folder.Value )
    {
        #Append folder number at start, construct full path
        [string]$newfolder = -join("0",$foldernumber,"_",$folder.Value)
        [string]$newfolder = -join($BASEFOLDER,'\',$newfolder)

        # Say what we do, do it
        Write-Output "[CREATE] folder: $newfolder"
        New-Item -ItemType Directory -Path $newfolder

        # Next folder get next number
        [int]$foldernumber = $foldernumber + 1 
    }
}







#======================================================================================================================================================================
#                      POSTPROCESSING                      =
#===========================================================

# PIN TO EXPLORER

echo "CheckIfCreateExplorerQuickAccess.Checked"
$CheckIfCreateExplorerQuickAccess.Checked

if ($true )
{
    Write-Output "[CREATE] Shortcut in File explorer"
    $o = new-object -com shell.application
    $o.Namespace($BASEFOLDER).Self.InvokeVerb("pintohome")
}

echo "$CheckIfCreateOutlookFolder.Checked"
$CheckIfCreateOutlookFolder.Checked
if ($true)
{
    Write-Output "[CREATE] Folder in Outlook"
    [string]$Username = $Env:UserName.split(".")[0]
    $TextInfo = (Get-Culture).TextInfo
    [string]$Username = $TextInfo.ToTitleCase($Username)
    [void]$ns.Folders.Item(1).Folders.Item("Posteingang").Folders.Item("02_ONGOING JOBS").Folders.Item($Username).Folders.Add($PROJECTNAME)
}



#==========================
#= INCLUDE ORIGINAL FILES =

# If user asked to include source files, include those in new folder, with naming conventions
if ($gui_filesource.SelectedItems.Text -notmatch $text_nofilesource)
{

    # CHECK WE HAVE THE MINIMUM FOLDERS
    # BECAUSE WE DONT KNOW WHAT TEMPLATE USER USED
    # IF THE STANDARD MINIMUM ISNT THERE, JUST USE BASE FOLDER INSTEAD
    [string]$INFO = "$BASEFOLDER\00_info"
    [string]$ORIG = "$BASEFOLDER\01_orig"

    if ($gui_filesource.SelectedItem -match $text_from_Outlook )
    {
        Write-Output "[DETECTED] Get source files from email"
        $sourcemail = $allgoodmails[$sourcefiles.SelectedItems.Index]
        foreach ($attachment in $sourcemail.Attachments)
        {
            if ($attachment.FileName -notmatch "^image[0-9][0-9][0-9]")
            {
                Write-Output (-join($ORIG,"\",$attachment.FileName))
                $attachment.SaveAsFile( -join($ORIG,"\",$attachment.FileName) )

            }
        } # End of attachment processing

    } # End of process outlook inclusion
    elseif ($gui_filesource.SelectedItem -match $text_from_Downloads )
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
    } # End of user load themselves


    # Before processing each source file,
    # Deal with the archives first
    Write-Output "Extracting all archives..."
    Get-ChildItem -Path $ORIG -Filter *.zip | Expand-Archive -DestinationPath $ORIG
    echo ok
    # PROCESS EACH SOURCE FILE
    # Rename and move file
    # Add count to total count and CSV
    # Ignore structure folders
    foreach ($file in (Get-ChildItem -Path "$ORIG" -Exclude "^[0-9][0-9]_" ))
    {
        echo "$ORIG"
        $newname = -join($DIRCODE,"_",$file.BaseName,"_orig",$file.Extension)
        Write-Output "[RENAME] As $newname"
        Rename-Item -Path $file.FullName -Newname "$newname"

    } # End of loop processing all source file

} # End of If we have source files



#==========================
#= START A TRADOS PROJECT =


# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.Checked)
{
	Write-Output "Starting Trados Studio..."
    # May not be where expected
    try {
        Set-Location "C:\Program Files (x86)\Trados\Trados Studio\Studio17"
        }
    catch { 
        $TRADOSDIR = (Get-ChildItem -Path "C:\Program Files (x86)" -Filter *.sdlproj -Recurse -ErrorAction SilentlyContinue -Force -File).Directory.FullName
        Write-Output "[DETECTED] Trados in $TRADOSDIR"
        Set-Location $TRADOSDIR
        }

    .\SDLTradosStudio.exe /createProject /name $PROJECTNAME
}


# OK NOW WE WORK
echo "CheckIfOpenExplorer.Checked"
$CheckIfOpenExplorer.Checked
if ($true )
{
    Write-Output "Starting Explorer..."
    start-process explorer "$BASEFOLDER"
}

echo "CheckIfOpenExplorer.Checked"
$CheckIfOpenExplorer.Checked

if ($CheckIfNotify.Checked )
{
    # Have a NICE NOTIFICATION THIS IS BALLERS
    # WOOOOHOOOO
    $objNotifyIcon                      = New-Object System.Windows.Forms.NotifyIcon
    #$objNotifyIcon.Icon = "M:\4_BE\06_General information\Stella\Skrivanek-Rocketlaunch\assets\Rocketlaunch-Icon.ico"  
    $objNotifyIcon.Icon                 = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
    $objNotifyIcon.BalloonTipTitle      = "Fertig!"
    $objNotifyIcon.BalloonTipIcon       = "Info"
#   $objNotifyIcon.BalloonTipText       = -join("Fertig !")
    $objNotifyIcon.Visible              = $True
    $objNotifyIcon.ShowBalloonTip(10000)
}