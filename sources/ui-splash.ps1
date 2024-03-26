


        #==============================================
        #                SPLASH Module                =
        #==============================================


#================================================================
# Create a splash. Just create it.

Write-Output "[START] Preparing Splash"

$UI_Splash                                      = New-Object System.Windows.Forms.Form
$UI_Splash.Text                                 = $APPNAME #-join($APPNAME," - ",$text_aboutsubtitle)
$UI_Splash.Width                                = 340
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
$UI_Splash_logo.Location                        = New-Object System.Drawing.Point(128,10)
$UI_Splash.Controls.Add($UI_Splash_logo)

$progressLabel                                  = New-Object System.Windows.Forms.Label
$progressLabel.Location                         = New-Object System.Drawing.Point(10,84)
$progressLabel.Size                             = New-Object System.Drawing.Size(320, 20)
$progressLabel.Text                             = $text_splash_loadingapp
$UI_Splash.Controls.Add($progressLabel)

$progressBar                                    = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location                           = New-Object System.Drawing.Point(10, 104)
$progressBar.Size                               = New-Object System.Drawing.Size(300, 20)
$progressBar.Value                              = 0
$UI_Splash.Controls.Add($progressBar)



