﻿

#================================================================
# Takes a string, change theme.

$UI_Splash                                      = New-Object System.Windows.Forms.Form
$UI_Splash.Text                                 = -join($APPNAME)
$UI_Splash.Width                                = 240
$UI_Splash.Height                               = 180
$UI_Splash.FormBorderStyle                      = "Fixed3D"
$UI_Splash.ControlBox                           = $false
$UI_Splash.StartPosition                        = 'CenterScreen'
$UI_Splash.Icon                                 = $icon

# FANCY ICON
$UI_Splash_logo                                 = new-object Windows.Forms.PictureBox
$UI_Splash_logo.Width                           = 64
$UI_Splash_logo.Height                          = 64
$UI_Splash_logo.Image                           = $image
$UI_Splash_logo.Location                        = New-Object System.Drawing.Point(78,10)
$UI_Splash.Controls.Add($UI_Splash_logo)

$progressLabel                                  = New-Object System.Windows.Forms.Label
$progressLabel.Location                         = New-Object System.Drawing.Point(10,84)
$progressLabel.Size                             = New-Object System.Drawing.Size(280, 20)
$progressLabel.Text                             = "0% Complete"
$UI_Splash.Controls.Add($progressLabel)

$progressBar                                    = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location                           = New-Object System.Drawing.Point(10, 104)
$progressBar.Size                               = New-Object System.Drawing.Size(200, 20)
#$progressBar.UseVisualStyleBackColor            = $true
$UI_Splash.Controls.Add($progressBar)



