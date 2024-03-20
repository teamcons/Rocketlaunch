

#================================================================
# Close app gracefully


function Change-Theme {
    param($selectedtheme)
    switch ($selectedtheme) {
        "Modern Color" {

            # Slick soft modern look
            $GUI_Form_MainWindow.BackColor          = "White"
            $panel_top.BackColor                    = "Orange"
            $panel_top.ForeColor                    = "Black"
            $Split.BackColor                        = "LightBlue"
            $bottom_panel.BackColor                 = '241,241,241'

            $panel_top.BackgroundImage.Dispose()
            $bottom_panel.BackgroundImage.Dispose()
        }
        "Brushed Metal" {

            # I liked that era of MacOs, it had some flair
            # Resizing gets very slow
            $GUI_Form_MainWindow.BackColor          = "White"
            $panel_top.BackColor                    = [System.Drawing.Color]::Transparent
            $panel_top.ForeColor                    = "Black"
            $Split.BackColor                        = "LightGray"
            $bottom_panel.BackColor                 = [System.Drawing.Color]::Transparent

            $panel_top.BackgroundImage              = [Drawing.Image]::FromFile(( -join($ScriptPath,'\assets\brushsteel.jpg')))
            $panel_top.BackgroundImageLayout        = "Stretch"
            $bottom_panel.BackgroundImage           = $panel_top.BackgroundImage
            $bottom_panel.BackgroundImageLayout     = "Stretch"

            $gui_okButton.UseVisualStyleBackColor       = $True
            $gui_cancelButton.UseVisualStyleBackColor   = $True
        }
        "Windows 98" {

            # Back to the good old classics
            # Need to disable VisualStyles too
            $GUI_Form_MainWindow.BackColor          = "Gray"
            $panel_top.BackColor                    = "Blue"
            $panel_top.ForeColor                    = "White"
            $Split.BackColor                        = "Gray"
            $bottom_panel.BackColor                 = 'Gray'

            $panel_top.BackgroundImage.Dispose()
            $bottom_panel.BackgroundImage.Dispose()

            $gui_okButton.UseVisualStyleBackColor       = $False
            $gui_cancelButton.UseVisualStyleBackColor   = $False
        }


        Default {}
    }
}


Change-Theme "Modern Color" #$THEME