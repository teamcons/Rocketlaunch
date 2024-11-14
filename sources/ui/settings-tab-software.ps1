
        #===================================================
        #                GUI - About Dialog                =
        #===================================================


$GUI_Tab_SoftwareSettings = New-object System.Windows.Forms.Tabpage
$GUI_Tab_SoftwareSettings.UseVisualStyleBackColor = $True 
$GUI_Tab_SoftwareSettings.Name = "About" 
$GUI_Tab_SoftwareSettings.Text = $text.Softwaresettings.tab




################################
# CHANGE LANGUAGE
$label_select_lang                     = New-Object System.Windows.Forms.Label
$label_select_lang.Text                = $text.Softwaresettings.lang
$label_select_lang.Top                 = 15
$label_select_lang.Left                = $GUI_Form_MainWindow_leftalign
$label_select_lang.Size                = New-Object System.Drawing.Size(200,20)

$combobox_select_lang                    = New-Object System.Windows.Forms.Combobox
$combobox_select_lang.Top                = ($label_select_lang.Top - 3)
$combobox_select_lang.Left               = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_lang.Size                = New-Object System.Drawing.Size(100,20)
$combobox_select_lang.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList


# For i in get-childiten localizations

Foreach ($language in (Get-ChildItem -Directory $MainDir/localization))
{
        [void]$combobox_select_lang.Items.Add($language.Name)
}

# Then default
$combobox_select_lang.SelectedItem = (Get-WinUserLanguageList).LanguageTag



# React to new select
$combobox_select_lang.Add_SelectedIndexChanged({
        $script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization -UICulture $combobox_select_lang.SelectedItem
})
    


################################################################
# CHANGE LANGUAGE
$label_select_theme                             = New-Object System.Windows.Forms.Label
$label_select_theme.Text                        = $text.Softwaresettings.theme
$label_select_theme.Top                         = $label_select_lang.Top + 35
$label_select_theme.Left                        = $GUI_Form_MainWindow_leftalign
$label_select_theme.Size                        = New-Object System.Drawing.Size(200,20)

$combobox_select_theme                          = New-Object System.Windows.Forms.Combobox
$combobox_select_theme.Top                      = ($label_select_theme.Top - 3)
$combobox_select_theme.Left                     = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_theme.Size                     = New-Object System.Drawing.Size(100,20)
$combobox_select_theme.DropDownStyle            = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$combobox_select_theme.Items.Add("Modern Color")
[void]$combobox_select_theme.Items.Add("Boring")
#[void]$combobox_select_theme.Items.Add("Brushed Metal")
[void]$combobox_select_theme.Items.Add("Windows 98")
[void]$combobox_select_theme.Items.Add("Princess Eyebleed")
$combobox_select_theme.SelectedItem = $combobox_select_theme.Items[0]



$combobox_select_theme.Add_SelectedIndexChanged({Change-Theme $combobox_select_theme.SelectedItem})
    

$GUI_Tab_SoftwareSettings.Controls.Add($label_select_lang)
$GUI_Tab_SoftwareSettings.Controls.Add($combobox_select_lang)
$GUI_Tab_SoftwareSettings.Controls.Add($label_select_theme)
$GUI_Tab_SoftwareSettings.Controls.Add($combobox_select_theme)

$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_SoftwareSettings)