





        #==========================================
        #                INTERNALS                =
        #==========================================


<#
All the gory innards, intermediate functions and all that
TODO: Split into two, have one dedicated to all the skrivanek-specific stuff

#>




#========================================
Write-Output "[START] Loading internal functions"


#================================================================
function load_template{
    param (  
        [System.Windows.Forms.DataGridView]$GRID,
        [string]$FILE)
    try {
        # Import all, skip the first one - It is the delimiter
        $detectedtemplate = (Import-Csv -Delimiter $SEP -Path $FILE -Header "Name","00","01","02","03","04","05","06","07","08","09" | Select-Object -Skip 1)
        foreach ($row in $detectedtemplate)
        {
            [void]$GRID.Rows.Add($row."Name",$row."00",$row."01",$row."02",$row."03",$row."04",$row."05",$row."06",$row."07",$row."08",$row."09",$row."10",$row."11",$row."12");
        }
    }
    catch {
        Write-Output "[ERROR] Cannot load templates, falling back to default"
        $GRID.Rows.Add("Minimal","info","orig");
        $GRID.Rows.Add("Standard TEP","info","orig","trados","to trans","from trans","to proof","from proof","to client")
    }
    return $GRID
}



#================================================================
function Get-CompanyName {
    param([string]$mailadress)
    $mailadress -match "@(?<content>).*"
    $attempt_at_companyname         = $matches[0].trim("@").split(".")[0]
    $attempt_at_companyname         = [cultureinfo]::GetCultureInfo("de-DE").TextInfo.ToTitleCase($attempt_at_companyname)
    return $attempt_at_companyname
}


#==========================================
# Try to predict what next number would be 
function Predict-StructCode {
 
    # Check if we can access the drive first, as it is a remote one
    # Try to go at the base structure
    try {Set-Location $ROOTSTRUCTURE}
    catch {
        $msg = -join("Struktur ist nicht verfügbar!`n",$ROOTSTRUCTURE," ist nicht erreichbar")
        [System.Windows.MessageBox]::Show($msg,"Nein",1)
        Close-All $GUI_Form_MainWindow
    }

    # Go at the latest project directory
    Set-Location (Get-ChildItem 2024_* -Directory | Select-Object -Last 1)

    # Retrieve latest project, clean year and project name out of i
    $PREDICT_CODE                               =  Get-ChildItem -Directory | Sort Name | Select-Object -Last 1
    $PREDICT_CODE                               =  $PREDICT_CODE.Name.Split("-")[1].Split("_")[0]
    [int]$PREDICT_CODE                          =  [int]$PREDICT_CODE + 1
    [bool]$script:CODE_PREDICTED                = $true
  
    # Nice
    return $PREDICT_CODE
    #$PREDICT_CODE = -join($YEAR,"-")
}





#================================================================
function Get-CleanifiedCodename {
    param([string]$PROJECTNAME)
    
    # Empty, so go on with what was initially predicted
    if ("$PROJECTNAME" -notmatch "^[0-9]" )
    {

        $PROJECTNAME = -join($PREDICT_CODE,$PROJECTNAME)
        Write-Output "Its words. Now: $PROJECTNAME"
    }

    # Remove invalid character, just in case
    $PROJECTNAME = $PROJECTNAME.Split([IO.Path]::GetInvalidFileNameChars()) -join ''
    #Write-Output "Removed invalid. Now: $PROJECTNAME"


    # is it missing zeros
    if ($PROJECTNAME -match "^[0-9][0-9][0-9]")
    {
        $PROJECTNAME = -join("0",$PROJECTNAME)
        #Write-Output "Missing first zero. Now: $PROJECTNAME"

    }
    elseif ($PROJECTNAME -match "^[0-9][0-9]")
    {
        $PROJECTNAME = -join("00",$PROJECTNAME)
        #Write-Output "Missing two zero. Now: $PROJECTNAME"
    }
    elseif ($PROJECTNAME -match "^[0-9]")
    {
        $PROJECTNAME = -join("000",$PROJECTNAME)
        #Write-Output "Missing three zero. Now: $PROJECTNAME"
    }

    # Add year
    $PROJECTNAME = -join($YEAR,"-",$PROJECTNAME)
    Write-Output "[INPUT] Got: $PROJECTNAME"
 
    return $PROJECTNAME
}






#================================================================
# Rebuilt the whole tree
# Take a valid project name
function Rebuild-Tree{
    param([string]$projectname)

    # Check if correct, as wrong input could be real bad
    Write-Output "[Rebuild-Tree] Rebuild whole tree - got $projectname"
    try { $DIRCODE = $projectname.SubString(0, 9) }
    catch {
        $ERRORTEXT="Projektcode ist unpassend !!!
    Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
    Angegeben: $PROJECTCODE"
        $btn = [System.Windows.Forms.MessageBoxButtons]::OK
        $ico = [System.Windows.Forms.MessageBoxIcon]::Information
        Add-Type -AssemblyName System.Windows.Forms 
        [void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)
    exit }


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
    return $BASEFOLDER

}



#================================================================
# Build the project at specified place
# Takes path to folder where to create, and an iterable with everything

function Create-AllFolders
{   param(
        [string]$basefolder,
        $tablewithfolders)

    # Count folder number
    [int]$foldernumber = 0

    foreach ($folder in $allfolderstocreate )
    {
        if ($folder.Value )
        {
            #Append folder number at start, construct full path
            [string]$newfolder = -join("0",$foldernumber,"_",($folder.Value.Split([IO.Path]::GetInvalidFileNameChars()) -join ''))
            [string]$newfolder = -join($BASEFOLDER,'\',$newfolder)

            # Say what we do, do it
            Write-Output "[CREATE] folder: $newfolder"
            New-Item -ItemType Directory -Path $newfolder

            # Next folder get next number
            [int]$foldernumber = $foldernumber + 1 
        }
    }



}




#================================================================
# Takes a directory
# Expands archives in it
# Rename all files with project code and orig
function Rename-Source
{
    param([string]$path,
            [string]$projectcode,
            [string]$orig)

    # Rename each file with code and "_orig"
    # Ignore structure folders
    foreach ($file in (Get-ChildItem -Path $path -Exclude "^[0-9][0-9]_" ))
    {
        $newname = -join($projectcode,"_",$file.BaseName,$orig,$file.Extension)
        Write-Output "[RENAME] As $newname"
        Rename-Item -Path $file.FullName -Newname "$newname"

    } # End of loop processing all source file


}




#================================================================
# Nuke datagridview
function Clear-Datagridview
{
    param($datagridview)

    $datagridview.Items.Clear()
    $datagridview.Columns.Clear()
    
}





#================================================================
# Send a notification. Yes, im used to Linux
# Take title and text
function Notify-Send
{
    param(
        [string]$title,
        [string]$text)
    
    Write-Output "[INFO] Notify $title $text"
    $objNotifyIcon                          = New-Object System.Windows.Forms.NotifyIcon
    $objNotifyIcon.Icon                     = $icon
    $objNotifyIcon.BalloonTipTitle          = $title
    $objNotifyIcon.BalloonTipIcon           = "Info"
    $objNotifyIcon.BalloonTipText           = $text
    $objNotifyIcon.Visible                  = $True
    $objNotifyIcon.ShowBalloonTip(5000)

    $objNotifyIcon.Visible                  = $False
    #$objNotifyIcon.Icon.Dispose();
    $objNotifyIcon.Dispose();

}




#================================================================
# Add a folder to File Explorer QuickAccess
# Takes a path to a folder
function Create-QuickAccess
{
    param([string]$folder)
    Write-Output "[CREATE] Shortcut in File explorer"
    $o = new-object -com shell.application
    $o.Namespace($folder).Self.InvokeVerb("pintohome")
}




#================================================================
function Start-TradosProject
{
    param([string]$projectname,[string]$projectfiles)

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

    .\SDLTradosStudio.exe /createProject /name $projectname #/files $projectfiles\*

}






#================================================================
# Close app gracefully
function Save-DataGridView
{
    param($datagridview,$file)
    
    # Just in case, check the variable hasnt been doestroyened
    if (($file -ne $none) -and ($file -ne ""))
    {
        # Create the CSV, specify separator to avoid issues opening the csv in your fav office software
        # Also no append, so we smonch the previous template
        Write-Output (-join("sep=",$SEP)) | Out-File -FilePath "$file"

        # Rebuild and append each line
        foreach ($row in $datagridview.Rows )
        {

            if (($row[0].Cells[0].Value -ne $none) -and ($row[0].Cells[0].Value -ne ""))
            {

                $rebuiltrow = ""
                for($i = 0; $i -lt 12; $i++) {
                    $rebuiltrow = -join($rebuiltrow, $row[0].Cells[$i].Value,$SEP)
                }

                Write-Output $rebuiltrow | Out-File -FilePath "$file" -Append

            }


        }
    } # end of if
} # end offunction









#================================================================
# Close app gracefully
function Close-All {
    param($GUI)

    Write-Output "[INPUT] Got Cancel. Aw. Exit."

    Try {$GUI.hide();}
    Catch {Write-Output "Oop"}
    #$GUI.Dispose();
    [System.Windows.Forms.Application]::Exit()
    
    #Stop-Process $pid
    exit
}









#================================================================
# Count all
function Count-AllWords {
    param($where, $saveto)

    # We, sadly, need Words
    $script:word            = New-Object -ComObject Word.Application
    [string]$analysisfile   = -join($saveto,"\",$default_csv_analysis)    
    [int]$totalcount        = 0    
    [float]$totaltime         = 0    

    # Create the CSV, specify separator to avoid issues opening the csv in your fav office software
    Write-Output (-join("sep=",$SEP)) | Out-File -FilePath $analysisfile

        # Add column headers
    $top = -join($text_csv_file,$SEP,$text_csv_wordcount,$SEP,$text_csv_proofreadtime,$SEP)
    Write-Output $top | Out-File -FilePath "$analysisfile" -Append 

    foreach ($file in (Get-ChildItem $where) )
    {

        # Use different backend depending on what needed
        # Each time, check the extension to know what we deal with
        if ($file.Extension -match ".doc[|x]" )
        {
            # OPEN IN WORD, PROCESS COUNT
            $filecontent = $word.Documents.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
            
        }
<#         elseif ($file.Extension -match ".xls[|x]" )
        {

            #foreach ($cell in $b.ActiveSheet.Rows[3].Cells) { if ($cell.Text -ne "") {$cell.Text} }

            # OPEN IN EXCEL, PROCESS COUNT
            $filecontent = $excel.Workbooks.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Excel.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
        }
        elseif ($file.Extension -match ".ppt[|x]" )
        {
            # OPEN IN POWRPOINT, PROCESS COUNT
            $filecontent = $powerpoint.Documents.Open($file.FullName)
            [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Powerpoint.WdStatistic]::wdStatisticWords)
            #CLOSE FILE
            $filecontent.Close()
        } #>
        elseif ($file.Extension -match ".pdf" )
        {
            # COUNT WORDS IN PDF FILE
            [int]$wordcount = (Get-Content $file.FullName | Measure-Object –Word).Words
        }
    
        elseif ($file.Extension -match ".[txt|csv|md|log]" )
        {
            # COUNT WORDS IN TXT FILE
            [int]$wordcount = (Get-Content $file.FullName | Measure-Object –Word).Words
        }
        else
        {
            # IDK
            [int]$wordcount = 0
        }
            
        # Update totalcount
        $proofreadtime      = [math]::round(($wordcount / $WORDS_PER_HOUR),$DECIMALS)
        $totalcount         = $totalcount + $wordcount
        $totaltime         = $totaltime + $proofreadtime
        $line               = -join($file.Name,$SEP,$wordcount,$SEP,$proofreadtime,$SEP)
        Write-Output $line | Out-File -FilePath $analysisfile -Append



	} # End of processing list
    
    $line = -join($text_csv_total,$SEP,$totalcount,$SEP,$totaltime,$SEP)
    Write-Output $line | Out-File -FilePath $analysisfile -Append

    Set-Clipboard $totalcount
    Import-Csv -Path $analysisfile `
        -Delimiter $SEP
        -Header $text_csv_file,$text_csv_wordcount,$text_csv_proofreadtime |
        Select-Object -Skip 1 |
        Out-GridView –Title "Rocketlaunch"

} # End of Count-Allwords