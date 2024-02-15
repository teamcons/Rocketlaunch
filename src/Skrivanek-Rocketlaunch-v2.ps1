#================================================================================================================================

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
Write-Output "Git: https://github.com/teamcons/Skrivanek-Rocketlaunch"
Write-Output ""
Write-Output ""


#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$APPNAME            = "-Rocketlaunch!"
[string]$PROJECTTEMPLATE    = "Minimal"

# Load templates from a csv in same place as executable
#[string]$LOAD_TEMPLATES_FROM = $MyInvocation.MyCommand.Path
[string]$ROOTSTRUCTURE      = "M:\9_JOBS_XTRF\"
[regex]$CODEPATTERN         = -join($YEAR,"-[0-9]")
[string]$YEAR               = get-date –f yyyy


#========================================
# Defaults

[string]$default_filesfrom          = "outlook"
[string]$default_fromdisk           = "$env:USERPROFILE\Downloads\"
[bool]$default_doanalysis           = $false
[bool]$default_opentrados           = $true
[bool]$default_createshortcut       = $true
[bool]$default_outlookfolder        = $true
[bool]$default_movesourcemail       = $true
[bool]$default_openexplorer         = $true


#========================================
# Localization

[string]$text_projectname           = "Projektname"
[string]$text_standardtab           = "Standard"
[string]$text_loadfilesfrom         = 'Email mit Ausgangsdatei'
[string]$text_usewhichtemplate      = 'Welche Projektvorlage soll verwendet werden?'
[string]$text_advancedtab           = "Erweitert"
[string]$text_doanalysis            = "Eine analyse machen ? (Langsam)"
[string]$text_opentrados            = "Trados öffnen ?"
[string]$text_openexplorer          = "Explorer öffnen ?"
[string]$text_help                  = "Hilfe"
[string]$text_OK                    = "Los!"
[string]$text_Cancel                = "Nö"



#========================================
# Outlook Capabilities

Write-Output "[STARTUP] Outlook Capabilities"
$OL                         = New-Object -ComObject OUTLOOK.APPLICATION
$ns                         = $OL.GETNAMESPACE("MAPI")
$date                       = Get-Date (Get-Date).AddDays(-2) -Format 'dd/MM/yyyy'
$filter                     = "[ReceivedTime] >= '$date 17:30'"
$allmails                   = $ns.Folders.Item(1).Folders.Item("Posteingang").Items.Restrict($filter)




# TRADOS. TODO: MORE FLEXIBLE
#[string]$NEWPROJECTICON     = "C:\Program Files (x86)\Trados\Trados Studio\Studio17\StudioTips\de\00_WelcomeFlow\NewProject.ico"



#==========================================
# Try to predict what next number would be 
# Catch: have at least first part of code
Write-Output "[STARTUP] Dircode prediction"
try
{
    # 
    Set-Location $ROOTSTRUCTURE
    Set-Location (Get-ChildItem 2024_* -Directory | Select-Object -Last 1)   
    $PREDICT_CODE           =  (Get-ChildItem -Directory | Select-Object -Last 1).Name.Substring(5,4)
    $PREDICT_CODE           =  [int]$PREDICT_CODE + 1
    [string]$PREDICT_CODE   =  -join($YEAR,"-",$PREDICT_CODE,"_")
    Write-Output "[PREDICTED] Next is $PREDICT_CODE"
}
catch
{
    $PREDICT_CODE = -join($YEAR,"-")
}


#================================
# Project icon in Base 64
# [Convert]::ToBase64String((Get-Content "..\assets\Skrivanek-Rocketlaunch-Icon.ico" -Encoding Byte))
Write-Output "[STARTUP] Loading icon"

[string]$iconBase64 = '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Gv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGs9jKxoWYysaAGMrGgBjKxoAYysaAGMrGkhjKxr2Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2EpGP+JVUv/3LWz/+C4tv/iu7n/7NjY/+zY2P/iu7n/4727/+jT0v+MYFf/YSgX/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr2YysaSGMrGgBjKxoAYysaAGMrGgBjKxqFYysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9iKh3/SB14/0ontv9aQNj/ZFLx/2dX9v9mV/b/ZFHw/2VS8f9kVPb/Vzeb/2IrHv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGoVjKxoAYysaAGMrGgBjKxoKYysau2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxn/YSoi/zgSk/8oB7r/OiXl/0M2//9DNf7/QzX+/0M1//9DNf//QzX//0w0xP9iKyL/YysZ/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxq7YysaCmMrGgBjKxoAYysaIWMrGt9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9iKhn/YikY/18oIP83EJH/KAa5/zkk5P9DNv//QzX+/0M1/v9DNf7/QzX+/0M1//9LMsH/YCkg/2EpF/9iKhn/Yysa/2MrGv9jKxr/Yysa/2MrGf9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa32MrGiFjKxoAYysaAGMrGj9jKxrzYysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2EqIf9iKhv/YysW/2MrF/9iKRf/g1NB/7OOff+zkoj/n4fB/5iD1f+gken/pZn2/6SZ9v+kmfb/pJn2/6SZ9v+kmff/qJjZ/7OUi/+ykYX/gVJD/2IpF/9jKxf/YysW/2MrG/9iKyL/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGvNjKxo/YysaAGMrGgBjKxpcYysa/WMrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/YysZ/1wnMP89FYn/SyeQ/1kwaP9dLUX/Yi84/8Smmf//8N////bs///57///+e////nu///47v//+O7///ju///47v//+O7///ju///57///+fH///vz/8Opo/9hLzf/XS1F/1kvZ/9RMp//TzO1/18tN/9jKxj/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr9YysaXGMrGgBjKxoAYysadmMrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGf9IG2j/LxDF/0Ex+P9GN/v/RjXt/2BN3v/l0tb//u/f///17f//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///47v/l2uj/X0zh/0Y17f9GN/r/RTf//0U3//9VMYb/YysY/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGnZjKxoAYysaAGMrGohjKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGf9fKSf/NhCY/zYd2f9FN/7/RTf+/0I0//+AcfT/+Oba//7w4///9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7/+vHv/35y+/9CNP//RTf+/0U3/v9FN///SzTN/2EsKv9jKxn/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxqIYysaAGMrGgBjKxqUYysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9kKxj/ViNC/y8Msv89Ken/RTj//0U3/v9ENv7/qZnq//7t2f/+8+j///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///47v+nm/b/RDb+/0U3/v9FN/7/RTf//0c28f9cLk//YysX/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/YysalGMrGgBjKxoAYysamGMrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/ZCwY/0weW/8uDr//QTDz/0U3//9FN/7/TD79/8u64///7tv///Tr///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//+e7/y8Dz/0s9/f9FN/7/RTf+/0U3/v9FN/3/VzBz/2MrFv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGphjKxoAYysaAGMrGpVjKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2QrGP9GGmv/LxHG/0Mz+P9FN/7/RDb+/1tN+v/k0t7//u/e///17f//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///ju/+Ta8P9aTPz/RDb+/0U3/v9FN/7/RTf//1QxjP9jKxf/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxqVYysaAGMrGgBjKxqMYysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Qxh1/zASyf9DNPr/RTf+/0M1/v9wYfb/8+Hb//7v4f//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///27v//9u7///bu///37v/06+//bmH6/0M1/v9FN/7/RTf+/0U3//9SMpj/YysZ/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/YysajGMrGgBjKxoAYysae2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/YysZ/0QZcf8vEsj/QzT5/0U3/v9CNP//hXbx//ro2f/+8eT///bu///27v//9u7///bu///27v//9u////fv///37///9u////bu///27v//9u7///bu///27v//9u7//PPu/4N3+f9CNP7/RTf+/0U3/v9FN///UzGU/2MrGP9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGntjKxoAYysaAGMrGmNjKxr+Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2QrGP9JG2T/Lg/D/0Iy9v9FN/7/QjX//5qK7f/97Nn//vLm///27v//9u7///bu///37///9+//++/m//Xm3P/15tz/++/m///37///9+////bu///27v//9u7///bu///37v+XjPf/QjT+/0U3/v9FN/7/RTf//1Uwgv9jKxf/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2MrGv9jKxr/Yysa/2Mr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'

# phew

# Rebuild an image from base 64
$iconBytes = [Convert]::FromBase64String($iconBase64)

# initialize a Memory stream holding the bytes
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

# This way we can draw icons without having any external file




#==============================================================
#                GUI - Ask the Right Questions                =
#==============================================================



#================
#= INITIAL WORK =

[int]$form_leftalign = 30
[int]$form_verticalalign = 470

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Lets look cool
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

$form                   = New-Object System.Windows.Forms.Form
$form.Text              = $APPNAME
$form.Size              = New-Object System.Drawing.Size(700,($form_verticalalign + 95))
#$form.MinimumSize       = New-Object System.Drawing.Size(600,450)
#$form.MaximumSize       = New-Object System.Drawing.Size(750,550)
$form.AutoSize          = $true
$form.AutoScale         = $true
$form.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)

$form.StartPosition     = 'CenterScreen'
#$form.FormBorderStyle  = 'FixedDialog'
$form.MaximizeBox       = $false
$form.Topmost           = $True
$form.BackColor        = "White"
$form.Icon              = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))



#==============
#= INPUT TEXT =

# FANCY ICON
$pictureBox             = new-object Windows.Forms.PictureBox
$pictureBox.Location    = New-Object System.Drawing.Point($form_leftalign,20)

#$img                    = [System.Drawing.Image]::Fromfile($NEWPROJECTICON);
$img = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$pictureBox.Width       = 64 #$img.Size.Width
$pictureBox.Height      = 64 #$img.Size.Height
$pictureBox.Image       = $img;
$form.controls.add($pictureBox)

# LABEL AND TEXT
# Label above input
$label                  = New-Object System.Windows.Forms.Label
$label.Location         = New-Object System.Drawing.Point(($form_leftalign + 80),30)
$label.Size             = New-Object System.Drawing.Size(170,20)
$label.AutoSize         = $true
$label.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Bold)
$label.Text             = 'Projektname:'
$form.Controls.Add($label)

# Input box
$textBox                = New-Object System.Windows.Forms.TextBox
$textBox.Location       = New-Object System.Drawing.Point(($form_leftalign + 80),60)
$textBox.Size           = New-Object System.Drawing.Size(170,30)
$textBox.Text           = $PREDICT_CODE
$form.Controls.Add($textBox)
$form.Add_Shown({$textBox.Select()})



#===================
#= SOURCE FILES    =


# Label above input
$labelsourcefiles                  = New-Object System.Windows.Forms.Label
$labelsourcefiles.Location         = New-Object System.Drawing.Point($form_leftalign,100)
$labelsourcefiles.Size             = New-Object System.Drawing.Size(240,20)
$labelsourcefiles.Text             = $text_loadfilesfrom
$form.Controls.Add($labelsourcefiles)


#'Windows icons
# for index: https://renenyffenegger.ch/development/Windows/PowerShell/examples/WinAPI/ExtractIconEx/imageres.html
# mail ,015
# folder ,013
# folder ,103
$imagelist                      = new-Object System.Windows.Forms.ImageList 
$imagelist.ImageSize            = New-Object System.Drawing.Size(16,16) # Size of the pictures
$icon_folder                      = [Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\Explorer.exe")
$icon_mail                      = [Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\Explorer.exe")
$icon_nope                      = [Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\Explorer.exe")

$imageList.Images.Add("IconMail",$icon_mail)

## Configure the ListView
$sourcefiles                   = New-Object System.Windows.Forms.ListView
$sourcefiles.Width             = 450 #$form.ClientRectangle.Width
$sourcefiles.Height            = 150 # $Form.ClientRectangle.Height
$sourcefiles.Location          = New-Object System.Drawing.Size($form_leftalign,120) 
$sourcefiles.Size              = New-Object System.Drawing.Size(570,150) 
$sourcefiles.FullRowSelect = $True
$sourcefiles.AutoResizeColumns(2)
$sourcefiles.View              = [System.Windows.Forms.View]::Details
$sourcefiles.SmallImageList = $imageList

#$sourcefiles.Columns.Add("Art")
$sourcefiles.Columns.Add("Betreff",360)
$sourcefiles.Columns.Add("Von",200)
#$sourcefiles.Columns.Add("Empfangen",100)
#$sourcefiles.Columns.Add("Dateien")


$allgoodmails               = New-Object System.Collections.ArrayList
[int]$goodmailindex = 0

# Look for emails with attachments
foreach ($mail in $allmails)
{

    [bool]$AddToGoodMails = $false
    foreach ( $attach in $mail.Attachments ) 
    {
        #echo $attach.FileName
        if ($attach.FileName -match  "(.pdf|.doc|.xls)" )
        {
            $AddToGoodMails = $true

        }   # End of each mails
    } #End of checking attachments

    # we found one with attachment !
    if ($AddToGoodMails -eq $true)
    {
        $allgoodmails.Add($mail)

        # Add to da list
        $sourcefilesItem = New-Object System.Windows.Forms.ListViewItem($mail.Subject)
        #$sourcefilesItem.Subitems.Add("mail")
        #$sourcefilesItem.Subitems.Add($windowsicons)
        [void]$sourcefilesItem.Subitems.Add($mail.SenderName)
        #$sourcefilesItem.Subitems.Add($Mail.ReceivedTime.ToString("hh:mm, u\m dd/MM"))
        #$sourcefilesItem.Subitems.Add("no")
        [void]$sourcefiles.Items.Add($sourcefilesItem)
        $goodmailindex += 1

    } # End of adding goodmail

} # End of looking for emails with attachments


# Add few other
#[void] $Combobox.Items.Add("(Ich möchte die Dateien selber holen)")
#[void] $Combobox.Items.Add("(Keine Ausgangsdatei, Danke!)")

$sourcefilesItem = New-Object System.Windows.Forms.ListViewItem("(Ich möchte die Dateien selber holen)")
#$sourcefilesItem.Subitems.Add($windowsicons,013)
#$sourcefilesItem.Subitems.Add("ordner")
[void]$sourcefiles.Items.Add($sourcefilesItem)

#$sourcefilesItem = New-Object System.Windows.Forms.ListViewItem("(Keine Ausgangsdatei, Danke!)")
#$sourcefilesItem.Subitems.Add($windowsicons,103)
#$sourcefilesItem.Subitems.Add("nein")
#$sourcefiles.Items.Add($sourcefilesItem)

#$sourcefiles.SelectedItem = $sourcefiles.Items[0]
#$Combobox.SelectedItem = $Combobox.Items[0]

# Maybe we already have a name
$allgoodmails[0].SenderEmailAddress -match "@(?<content>.*).com"
$attempt_at_companyname         = $matches["content"]
$attempt_at_companyname         = [cultureinfo]::GetCultureInfo("de-DE").TextInfo.ToTitleCase($attempt_at_companyname)
$textBox.Text                   = -join($PREDICT_CODE,$attempt_at_companyname)


$sourcefiles.Items.Add("IconMail",0)


# Add the ListView to the Form
$form.Controls.Add($sourcefiles)



#=====================
#= LIST OF TEMPLATES =

# $labeltemplate                  = New-Object System.Windows.Forms.Label
# $labeltemplate.Location         = New-Object System.Drawing.Point($form_leftalign,260)
# $labeltemplate.Size             = New-Object System.Drawing.Size(215,30)
# $labeltemplate.Text             = 'Welche Projektvorlage soll verwendet werden?'
# $form.Controls.Add($labeltemplate)
# $templates                        = New-Object System.Windows.Forms.ListView
# $templates.Location               = New-Object System.Drawing.Point($form_leftalign,290)
# $templates.Size                   = New-Object System.Drawing.Size(450,200)
# $templates.AutoSize               = $true 
# $templates.Height                 = 100
# $templates.FullRowSelect = $True
# $templates.AutoResizeColumns(2)
# $templates.View              = [System.Windows.Forms.View]::Details
# $templates.Columns.Add("Vorlage",100)
# $templates.Columns.Add("Ordner",250)
# ## LOAD FROM CSV HERE
# $templatesItem = New-Object System.Windows.Forms.ListViewItem("Minimal")
# [void]$templatesItem.Subitems.Add("00_info, 01_orig")
# [void]$templates.Items.Add($templatesItem)
# $form.Controls.Add($templates)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point($form_leftalign,300)
$label2.Size = New-Object System.Drawing.Size(500,30)
$label2.Text = $text_usewhichtemplate
$form.Controls.Add($label2)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point($form_leftalign,330)
$listBox.Size = New-Object System.Drawing.Size(670,120)
$listBox.Height = 120

[void] $listBox.Items.Add('Minimal')
$listBox.SelectedItem = "Minimal"

## LOAD FROM CSV HERE
[void] $listBox.Items.Add('Standard TEP')
[void] $listBox.Items.Add('Provider macht TEP')
[void] $listBox.Items.Add('Sworn Translation')
[void] $listBox.Items.Add('MemoQ')
[void] $listBox.Items.Add('Production')

$form.Controls.Add($listBox)





#===========
#= OPTIONS =


# Check if include original files
#$CheckIfSourceFiles = New-Object System.Windows.Forms.CheckBox        
#$CheckIfSourceFiles.Location = New-Object System.Drawing.Point(20,240)
#$CheckIfSourceFiles.Size = New-Object System.Drawing.Size(215,25)
#$CheckIfSourceFiles.Text = "Ausgangsdateien einbeziehen?"
#$CheckIfSourceFiles.UseVisualStyleBackColor = $True
#$CheckIfSourceFiles.Checked = $True
#$form.Controls.Add($CheckIfSourceFiles)


# Check if count words ?
#$CheckIfAnalysis                = New-Object System.Windows.Forms.CheckBox        
#$CheckIfAnalysis.Location       = New-Object System.Drawing.Point(($form_leftalign + 300),300)
#$CheckIfAnalysis.Size           = New-Object System.Drawing.Size(250,25)
#$CheckIfAnalysis.Text           = "Wortzahl machen? (Langsam)"
#$CheckIfAnalysis.UseVisualStyleBackColor = $True
#$CheckIfAnalysis.Checked        = $True
#$form.Controls.Add($CheckIfAnalysis)



# Check if start new trados project
#$CheckIfTrados                  = New-Object System.Windows.Forms.CheckBox        
#$CheckIfTrados.Location         = New-Object System.Drawing.Point(($form_leftalign + 300),330)
#$CheckIfTrados.Size             = New-Object System.Drawing.Size(250,25)
#$CheckIfTrados.Text             = "Ein neues Trados-Projekt beginnen?"
#$CheckIfTrados.UseVisualStyleBackColor = $True
#$CheckIfTrados.Checked          = $True
#$form.Controls.Add($CheckIfTrados)



#====================
#= OKCANCEL BUTTONS =


$gui_advanced_settings                   = New-Object System.Windows.Forms.Button
$gui_advanced_settings.Location          = New-Object System.Drawing.Point(($form_leftalign),$form_verticalalign)
$gui_advanced_settings.Size              = New-Object System.Drawing.Size(120,30)
$gui_advanced_settings.Text              = 'Erweitert...'
$gui_advanced_settings.UseVisualStyleBackColor = $True
$gui_advanced_settings.add_click({
    [System.Windows.Forms.MessageBox]::Show("Hello World. " + ($firstNameTextBox.Text) + " " + ($lastNameTextbox.Text) , "My Dialog Box")
})


$form.Controls.Add($gui_advanced_settings)



$gui_okButton                   = New-Object System.Windows.Forms.Button
$gui_okButton.Location          = New-Object System.Drawing.Point(($form_leftalign + 400),$form_verticalalign)
$gui_okButton.Size              = New-Object System.Drawing.Size(120,30)
$gui_okButton.Text              = 'Los!'
$gui_okButton.UseVisualStyleBackColor = $True
$gui_okButton.DialogResult      = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton          = $gui_okButton
#$gui_okButton.BackColor =”Green”
$form.Controls.Add($gui_okButton)

$gui_cancelButton               = New-Object System.Windows.Forms.Button
$gui_cancelButton.Location      = New-Object System.Drawing.Point(($form_leftalign + 550),$form_verticalalign)
$gui_cancelButton.Size          = New-Object System.Drawing.Size(120,30)
$gui_cancelButton.Text          = 'Nö'
$gui_cancelButton.UseVisualStyleBackColor = $False #$True
$gui_cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton          = $gui_cancelButton
$form.Controls.Add($gui_cancelButton)



#==============
#= WRAP IT UP =


$result = $form.ShowDialog()

[string]$PROJECTNAME        = $textBox.Text 
$PROJECTTEMPLATE            = $listBox.SelectedItem

# Cancel culture
# Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Output "[INPUT] Got Cancel. Aw. Exit."
        exit
    }
Write-Output "[INPUT] Got: $PROJECTNAME"




#==============================================================
#                     Processing Le input                     =
#==============================================================


# Empty, so go on with what was initially predicted
if ("$PROJECTNAME" -notmatch "[0-9]" )
{

    $PROJECTNAME = -join($PREDICT_CODE,$PROJECTNAME)
    Write-Output "Its words. Now: $PROJECTNAME"
}

# Remove invalid character, just in case
$PROJECTNAME = $PROJECTNAME.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
Write-Output "Removed invalid. Now: $PROJECTNAME"



# Code is formed correctly buuuut...
# Person started with invalid char - the slash from XTRF or some shit
# Test for underscore because invalid charactersvebeen stripped
if ($PROJECTNAME -match "^20[0-9][0-9]_[0-9]")
{
    [regex]$pattern = '_'
    $PROJECTNAME = $pattern.replace($PROJECTNAME, '-', 1)
    Write-Output "Slash, replace with dash. Now: $PROJECTNAME"
}





# If it matches the pattern
# We'll probably just have to insert missing zeros, if
if ($PROJECTNAME -match "$CODEPATTERN" )
{

    Write-Output "Case : Wellformed code"
    

    # just check if it is missing zeros
    if ($PROJECTNAME -match "^20[0-9][0-9]-[0-9][0-9][0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"0")
        Write-Output "Missing first zero. Now: $PROJECTNAME"

    }
    elseif ($PROJECTNAME -match "^20[0-9][0-9]\-[0-9][0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"00")
        Write-Output "Missing two zero. Now: $PROJECTNAME"
    }
    elseif ($PROJECTNAME -match "^20[0-9][0-9]\-[0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"000")
        Write-Output "Missing three zero. Now: $PROJECTNAME"
    }


}
else { Write-Output "$PROJECTNAME does not match $CODEPATTERN" }

# does NOT Start with year, but with numbers, or slash-numbers.
if ($PROJECTNAME -notmatch "^[0-9][0-9][0-9][0-9]-")
{
    Write-Output "Case : No year, adding year"

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


}


###########################DEBUG
#Write-Output "DEBUG"
#Write-Output "IS CORRECT ?"
#Write-Output "$PROJECTNAME"
#exit
###########################DEBUG


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




## ULTIMATE ULTIMATE CHECK
if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and ($DIRCODE -match "^20[0-9][0-9]\-[0-9][0-9][0-9][0-9]" ))
{

Write-Output "[INPUT] Project Name is $PROJECTNAME"
Write-Output "[INPUT] Project Template is $PROJECTTEMPLATE"

Write-Output "[INPUT] Include Source Files ?" $CheckIfSourceFiles.CheckState
Write-Output "[INPUT] Start new Trados project ?" $CheckIfTrados.CheckState

Write-Output "[DETECTED] Dircode is $DIRCODE"

}
else { 	$ERRORTEXT="Projektcode oder vorlage ist unpassend !!!
Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
Angegeben: $PROJECTCODE"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)

exit  }




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


switch ( $PROJECTTEMPLATE )
{
    "Minimal"
        { $FOLDERS = @("00_info", "01_orig" ) }
    "Standard TEP"
        { $FOLDERS = @("00_info", "01_orig", "02_trados", "03_to trans", "04_from trans", "05_to proof", "06_from proof", "07_to client") }
    "Provider macht TEP"
        { $FOLDERS = @("00_info", "01_orig", "02_trados", "03_to TEP", "04_from TEP", "05_to client") }
    "Sworn Translation"
        { $FOLDERS = @("00_info", "01_orig", "02_to client") }
    "MemoQ"
        { $FOLDERS = @("00_info", "01_orig", "02_memoQ", "03_to client") }
    "Production"
        { $FOLDERS = @("00_info") }
}


# CREATE ALLLLL THE FOLDERS
foreach ($folder in $FOLDERS)
{
    Write-Output "[CREATE] folder: $BASEFOLDER\$folder"
    New-Item -ItemType Directory -Path "$BASEFOLDER\$folder"
}




# PIN TO EXPLORER
$o = new-object -com shell.application
$o.Namespace($BASEFOLDER).Self.InvokeVerb("pintohome")







#==================================================
#                      BONUS                      =
#==================================================



#==========================
#= INCLUDE ORIGINAL FILES =


# If user asked to include source files, include those in new folder, with naming conventions
if ($sourcefiles.SelectedItem -isnot "(Keine Ausgangsdatei, Danke!)") #($CheckIfSourceFiles.CheckState.ToString() -eq "Checked")
{


    if ($sourcefiles.SelectedItem -is "(Ich möchte die Dateien selber holen)") #($CheckIfSourceFiles.CheckState.ToString() -eq "Checked")
    {
        Write-Output "[DETECTED] Load source files"

        # Grab source files
        $SOURCEFILES = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            InitialDirectory    = $default_fromdisk
            Multiselect         = $true
            Title               = $APPNAME
        }
        $null = $SOURCEFILES.ShowDialog()
        Write-Output "[INPUT] Got:"
        Write-Output $SOURCEFILES.FileNames
    } # End of user load themselves
    else
    {
        
        Write-Output "[DETECTED] Get source files from email"
    }









    # CHECK WE HAVE THE MINIMUM FOLDERS
    # BECAUSE WE DONT KNOW WHAT TEMPLATE USER USED
    # IF THE STANDARD MINIMUM ISNT THERE, JUST USE BASE FOLDER INSTEAD
    if (Test-Path "$BASEFOLDER\00_info\" -PathType Container)
    {
            [string]$INFO = "$BASEFOLDER\00_info\"
    }
    else
    {
            [string]$INFO = "$BASEFOLDER\"
    }

    if (Test-Path "$BASEFOLDER\01_orig\" -PathType Container)
    {
            [string]$ORIG = "$BASEFOLDER\01_orig\"
    }
    else
    {
            [string]$ORIG = "$BASEFOLDER\"
    }


    # ONLY IF ANALYSIS WISHED
    if ($CheckIfAnalysis.CheckState.ToString() -eq "Checked")
    {

        # Need to use Word
        $word           = New-Object -ComObject Word.Application 
        $excel          = New-Object -ComObject Excel.Application 
        $powerpoint     = New-Object -ComObject Powerpoint.Application 
        $word.Visible   = $false 
        $excel.Visible  = $false 
        $powerpoint.Visible = $false 
        [int]$totalcount = 0
  
        # Create the CSV
        $ANALYSIS = -join($DIRCODE,"_","Analyse.csv")
        Write-Output "sep=;" | Out-File -FilePath "$INFO\$ANALYSIS"
        Write-Output "Datei;Wörterzahl" | Out-File -FilePath "$INFO\$ANALYSIS" -Append 
        



     }    


    # PROCESS EACH SOURCE FILE
    # Rename and move file
    # Add count to total count and CSV
    foreach ($file in $SOURCEFILES.FileNames)
    {



        if ($Combobox.SelectedItem -is "(Ich möchte die Dateien selber holen)") #($CheckIfSourceFiles.CheckState.ToString() -eq "Checked")
        {
            # DO THE MOVE
            $truefile = Get-Item "$file"
            $newname = -join($DIRCODE,"_",$truefile.BaseName,"_orig",$truefile.Extension)
            Write-Output "[MOVE] Move to $ORIG\$newname"
            Move-Item -Path "$truefile" -Destination "$ORIG\$newname"
        }
        else {
            # SAVE FROM OUTLOOK
            Write-Output "TODO TODO SAVE FROM OUTLOOK"
        }

        
        # ONLY IF ANALYSIS WISHED
        if ($CheckIfAnalysis.CheckState.ToString() -eq "Checked")
        {

            # Use different backend depending on what needed
            # Each time, check the extension to know what we deal with
            if ("$newname" -match ".[doc|docx]$" )
            {
                # OPEN IN WORD, PROCESS COUNT
                $filecontent = $word.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".[xls|xlsx]$" )
            {
                # OPEN IN EXCEL, PROCESS COUNT
                $filecontent = $excel.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Excel.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".[ppt|pptx]$" )
            {
                # OPEN IN POWRPOINT, PROCESS COUNT
                $filecontent = $powerpoint.Documents.Open("$ORIG\$newname")
                [int]$wordcount = $filecontent.ComputeStatistics([Microsoft.Office.Interop.Powerpoint.WdStatistic]::wdStatisticWords)
            }
            elseif ("$newname" -match ".pdf$" )
            {
                # COUNT WORDS IN PDF FILE
                [int]$wordcount = (Get-Content "$ORIG\$newname" | Measure-Object –Word).Words
            }

            elseif ("$newname" -match ".[txt|csv]$" )
            {
                # COUNT WORDS IN TXT FILE
                [int]$wordcount = (Get-Content "$ORIG\$newname" | Measure-Object –Word).Words
            }

            else
            {
                # IDK
                [int]$wordcount = 0
            }
        
            # USE THE WORDCOUNT
            [int]$totalcount += $wordcount
            Write-Output "Wordcount: $wordcount"
            Write-Output "$newname;$wordcount" | Out-File -FilePath "$INFO\$ANALYSIS" -Append 


            #CLOSE FILE
            $filecontent.Close()
        }


    } # End of loop processing all source file

    
    
    # ONLY IF ANALYSIS WISHED
    if ($CheckIfAnalysis.CheckState.ToString() -eq "Checked")
    {
        # Wont need Word anymore
        $word.Quit()
        $excel.Quit()
        $powerpoint.Quit()
    
        # Finish CSV file, 
        Write-Output "SUMME;$totalcount" | Out-File -FilePath "$INFO\$ANALYSIS" -Append

        # and create shortcut to it in orig for quick access
        # name has the totalcount for quicker overview
        Write-Output "[CREATE] Shortcute to Analysis"
        shortcutname    = -join("$totalcount","w","lnk")
        $WshShell       = New-Object -comObject WScript.Shell
        $Shortcut       = $WshShell.CreateShortcut("$ORIG\$shortcutname")
        $Shortcut.TargetPath = "$INFO\$ANALYSIS"
        $Shortcut.Save()


        # Clipboard
        Set-Clipboard -Value $totalcount
        Write-Output "[ACTION] Set clipboard to $totalcount"

   
        # Have a NICE NOTIFICATION THIS IS BALLERS
        # WOOOOHOOOO
        $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
        #$objNotifyIcon.Icon = "M:\4_BE\06_General information\Stella\Skrivanek-Rocketlaunch\assets\Rocketlaunch-Icon.ico"
        $objNotifyIcon.Icon             = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
        $objNotifyIcon.BalloonTipIcon   = "Info"
        $objNotifyIcon.BalloonTipText   = "Die Wortzahl ($totalcount) können Sie über Strng+V einfügen ;)"
        $objNotifyIcon.BalloonTipTitle  = "Wortzahl Zur Zwischenablage hinzugefügt!"
        $objNotifyIcon.Visible          = $True
        $objNotifyIcon.ShowBalloonTip(10000)
    } # End of Cleanup analysis
   
    


} # End of If we have source files



#==========================
#= START A TRADOS PROJECT =


# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.CheckState.ToString() -eq "Checked")
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


#=============
#= LAST STEP =

# OK NOW WE WORK
Write-Output "Starting Explorer..."
start-process explorer "$BASEFOLDER"

