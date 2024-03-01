#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1"
# Usage: Compiled form, just double-click.



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
# Grab script location in a way that is compatible with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }


    


#========================================
# Get all resources
Write-Output "[START] Loading text"
Import-Module $ScriptPath/text.ps1
Write-Output "[START] Loading defaults"

Import-Module $ScriptPath/defaults.ps1
Write-Output "[START] Loading icon"

Import-Module $ScriptPath/bigstring.ps1
Write-Output "[START] Loading internal functions"

Import-Module $ScriptPath/internals.ps1

Write-Output "[START] Loading graphical user interface"
Import-Module $ScriptPath/ui.ps1 

Write-Output "[START] Loading Outlook capabilities"
Import-Module $ScriptPath/outlook-backend.ps1 



    #==============================================================
    #                                                             =
    #                     Processing Le input                     =
    #                                                             =
    #==============================================================

Write-Output "[START] Show main window"
$result = $GUI_Form_MainWindow.ShowDialog()

# Cancel culture : Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    { Write-Output "[INPUT] Got Cancel. Aw. Exit." ; exit }


[string]$PROJECTNAME        = $gui_code.Text 
Write-Output "[INPUT] Got: $PROJECTNAME"


# Make sure we have clean input
[string]$PROJECTNAME                = (Get-CleanifiedCodename $PROJECTNAME)[-1]





[string]$BASEFOLDER = Rebuild-Tree $PROJECTNAME



#====================================================================================================
echo $PROJECTNAME
echo $BASEFOLDER
exit
#====================================================================================================


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

Create-QuickAccess $BASEFOLDER
Create-OutlookFolder $PROJECTNAME $ns

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






# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.Checked)
{
    Start-TradosProject $PROJECTNAME
}

# OK NOW WE WORK
start-process explorer "$BASEFOLDER"




