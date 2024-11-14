


        #=============================================
        #                THEME Module                =
        #=============================================



<# 
Play with interface stuff, see what looks good
See what looks tacky
Switcher works, but hidden from settings, because i feel bad about it
#>



#================================================================
# Takes a string, change theme.


function Change-Theme {
    param($selectedtheme)
    switch ($selectedtheme) {
        "Modern Colors" {

            # Slick soft modern look
            $GUI_Form_MainWindow.BackColor                          = "White"
            $GUI_Form_MoreStuff.BackColor                           = '237,237,237'

            # MAIN UI
            $panel_top.BackColor                                    = "Orange"
            $panel_top.ForeColor                                    = "Black"

            $panel_sourcefile.BackColor                             = "White"
            #$sourcefiles.AlternatingRowsDefaultCellStyle.BackColor = "LightBlue"
            #$sourcefiles.ColumnHeadersDefaultCellStyle.BackColor = "LightBlue"
            $Split.BackColor                                        = "LightBlue"
            $panel_template.BackColor                               = "White" #'Red'
            $templates.Backgroundcolor                              = "White"
            $templates.AlternatingRowsDefaultCellStyle.BackColor    = '237,237,237'


            $bottom_panel.BackColor                                 = '237,237,237'
<# 
            
            $gui_okButton.BackColor                     = "Green"
            $gui_okButton.ForeColor                     = "White"
            $gui_cancelButton.BackColor                  = "Red"
            $gui_cancelButton.ForeColor                  = "White" #>
            

            # If there is a theme with background images - dispose it
            #try {$panel_top.BackgroundImage.Dispose()}
            #catch {Write-Output "Nothing to dispose"}
            #try {$bottom_panel.BackgroundImage.Dispose()}
            #catch {Write-Output "Nothing to dispose"}


        }
        "Boring" {
            # No fancy. All defaults. Bleh.
            $GUI_Form_MainWindow.ResetBackColor()
            $GUI_Form_MoreStuff.ResetBackColor()
            
            $panel_top.ResetBackColor()
            $panel_top.ResetForeColor()
            $panel_top.BackgroundImage.Dispose()

            $panel_sourcefile.ResetBackColor()
            $Split.ResetBackColor()
            $templates.ResetBackColor()
            $panel_template.ResetBackColor()

            $bottom_panel.ResetBackColor()
            $bottom_panel.BackgroundImage.Dispose()
        }
<#         "Brushed Metal" {

            # I liked that era of MacOs, it had some flair
            # Resizing gets very slow
            $GUI_Form_MainWindow.BackColor          = "LightGray"
            $GUI_Form_MoreStuff.BackColor           = "LightGray"
 
            $panel_top.BackColor                    = [System.Drawing.Color]::Transparent
            $panel_top.ForeColor                    = "Black"
            $panel_top.BackgroundImage              = [Drawing.Image]::FromFile(( -join($ScriptPath,'\assets\brushsteel.jpg')))
            $panel_top.BackgroundImageLayout        = "Stretch"
 
            $panel_sourcefile.BackColor             = "LightGray"
            $Split.BackColor                        = "Blue"
            $templates.Backgroundcolor                    = "White"
            $panel_template.BackColor               = "LightGray" #'Red'

            $bottom_panel.BackColor                 = [System.Drawing.Color]::Transparent
            $bottom_panel.BackgroundImage           = $panel_top.BackgroundImage
            $bottom_panel.BackgroundImageLayout     = "Stretch"
        } #>
        "Windows 98" {

            # Back to the good old classics
            # Need to disable VisualStyles too
            $GUI_Form_MainWindow.BackColor          = "LightGray"
            $GUI_Form_MoreStuff.BackColor           = "LightGray"
            
            $panel_top.BackColor                    = "Blue"
            $panel_top.ForeColor                    = "White"
            $panel_top.BackgroundImage.Dispose()

            $panel_sourcefile.BackColor             = "LightGray"
            $Split.BackColor                        = "Blue"
            $templates.Backgroundcolor                    = "White"
            $panel_template.BackColor               = "LightGray" #'Red'

            $bottom_panel.BackColor                 = "LightGray"
            $bottom_panel.BackgroundImage.Dispose()
        }
        "Princess Eyebleed" {

            # MY EYES
            $GUI_Form_MainWindow.BackColor          = "Cyan"
            $GUI_Form_MoreStuff.BackColor           = "Cyan"
            
            $panel_top.BackColor                    = "Pink"
            $panel_top.ForeColor                    = "Yellow"
            $panel_top.BackgroundImage.Dispose()

            $panel_sourcefile.BackColor             = "Aqua"
            $Split.BackColor                        = "Orange"
            $templates.Backgroundcolor                    = "White"
            $panel_template.BackColor               = "Aqua" #'Red'

            $bottom_panel.BackColor                 = 'Pink'
            $bottom_panel.BackgroundImage.Dispose()

        }
    }
}


Change-Theme $THEME