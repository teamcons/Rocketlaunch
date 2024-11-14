
        #===================================================
        #                GUI - About Dialog                =
        #===================================================


$GUI_Tab_SoftwareSettings = New-object System.Windows.Forms.Tabpage
$GUI_Tab_SoftwareSettings.UseVisualStyleBackColor = $True 
$GUI_Tab_SoftwareSettings.Name = "About" 
$GUI_Tab_SoftwareSettings.Text = $text.Softwaresettings.tab




################################
# CHANGE LANGUAGE
$label_select_box                     = New-Object System.Windows.Forms.Label
$label_select_box.Text                = $text.Softwaresettings.box
$label_select_box.Top                 = 15
$label_select_box.Left                = $GUI_Form_MainWindow_leftalign
$label_select_box.Size                = New-Object System.Drawing.Size(200,20)

$script:combobox_select_box                    = New-Object System.Windows.Forms.Combobox
$combobox_select_box.Top                = ($label_select_box.Top - 3)
$combobox_select_box.Left               = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_box.Size                = New-Object System.Drawing.Size(100,20)
$combobox_select_box.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList




################################
# CHANGE LANGUAGE
$label_select_lang                     = New-Object System.Windows.Forms.Label
$label_select_lang.Text                = $text.Softwaresettings.lang
$label_select_lang.Top                 = $label_select_box.Top + 35
$label_select_lang.Left                = $GUI_Form_MainWindow_leftalign
$label_select_lang.Size                = New-Object System.Drawing.Size(200,20)

$combobox_select_lang                    = New-Object System.Windows.Forms.Combobox
$combobox_select_lang.Top                = ($label_select_lang.Top - 3)
$combobox_select_lang.Left               = ($GUI_Form_MainWindow_leftalign + 200 )
$combobox_select_lang.Size                = New-Object System.Drawing.Size(100,20)
$combobox_select_lang.DropDownStyle           = [System.Windows.Forms.ComboBoxStyle]::DropDownList


# For i in get-childiten localizations

[void]$combobox_select_lang.Items.Add($text.Softwaresettings.langdefault)
Foreach ($language in (Get-ChildItem -Directory $MainDir/localization))
{
        [void]$combobox_select_lang.Items.Add($language.Name)
}

# Then default
$combobox_select_lang.SelectedItem = $settings.Preferences.Language
$script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization


# React to new select
$combobox_select_lang.Add_SelectedIndexChanged({

        # If it is default, revert to system language
        if ($combobox_select_lang.SelectedItem -match $text.Softwaresettings.langdefault)
        {
                $script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization
        }
        # Else take whatever is indicated
        else {
                $script:text = Import-LocalizedData -FileName interface.psd1 -BaseDirectory $MainDir\localization -UICulture $combobox_select_lang.SelectedItem
        }
        # Save choice in settings
        $settings.UI.Language = $combobox_select_lang.SelectedItem
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
[void]$combobox_select_theme.Items.Add("Modern Colors")
[void]$combobox_select_theme.Items.Add("Boring")
#[void]$combobox_select_theme.Items.Add("Brushed Metal")
[void]$combobox_select_theme.Items.Add("Windows 98")
[void]$combobox_select_theme.Items.Add("Princess Eyebleed")
$combobox_select_theme.SelectedItem = $combobox_select_theme.Items[0]


################################################################
# Save
$combobox_select_theme.Add_SelectedIndexChanged({Change-Theme $combobox_select_theme.SelectedItem})
    




################################################################
# Stitch back

$GUI_Settingstabsoftware_Closebutton                               = New-Object System.Windows.Forms.Button
$GUI_Settingstabsoftware_Closebutton.Text                          = $text.Settings.close
$GUI_Settingstabsoftware_Closebutton.Size                          = New-Object System.Drawing.Size(100,25)
$GUI_Settingstabsoftware_Closebutton.Left                          = 215
$GUI_Settingstabsoftware_Closebutton.Top                           = $buttonalign
$GUI_Settingstabsoftware_Closebutton.Add_Click( {Save-Settings ; $GUI_Form_MoreStuff.Close() } )



$GUI_Tab_SoftwareSettings.Controls.Add($label_select_box)
$GUI_Tab_SoftwareSettings.Controls.Add($combobox_select_box)

$GUI_Tab_SoftwareSettings.Controls.Add($label_select_lang)
$GUI_Tab_SoftwareSettings.Controls.Add($combobox_select_lang)
$GUI_Tab_SoftwareSettings.Controls.Add($label_select_theme)
$GUI_Tab_SoftwareSettings.Controls.Add($combobox_select_theme)

$GUI_Tab_SoftwareSettings.Controls.Add($GUI_Settingstabsoftware_Closebutton)


$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_SoftwareSettings)



