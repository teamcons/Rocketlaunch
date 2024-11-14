
        #===================================================
        #                GUI - About Dialog                =
        #===================================================


$GUI_Tab_About                  = New-object System.Windows.Forms.Tabpage
$GUI_Tab_About.UseVisualStyleBackColor   = $True 
$GUI_Tab_About.Name             = "About" 
$GUI_Tab_About.Text             = $text.About.tab


# FANCY ICON
$applogo                        = new-object Windows.Forms.PictureBox
$applogo.Width                  = 64
$applogo.Height                 = $applogo.Width
$applogo.Image                  = $image
$applogo.Location               = New-Object System.Drawing.Point(128,20)

#$img                    = (get-item $ScriptPath\assets\icon-mini.ico)
#$pictureBox.Image       = [system.drawing.image]::FromFile($img)

# Label above input
$abouttitle                     = New-Object System.Windows.Forms.Label
$abouttitle.Text                = "-Rocketlaunch!"
$abouttitle.Size                = New-Object System.Drawing.Size(280,20)
$abouttitle.Left                = ($GUI_Form_MainWindow_leftalign + 85)
$abouttitle.Top                 = 95
$abouttitle.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif', 13, [System.Drawing.FontStyle]::Bold)

# Label above input
$aboutsubtitle                  = New-Object System.Windows.Forms.Label
$aboutsubtitle.Text             = $text.About.tagline
$aboutsubtitle.Size             = New-Object System.Drawing.Size(360,20)
$aboutsubtitle.Left             = ($GUI_Form_MainWindow_leftalign + 40)
$aboutsubtitle.Top              = 120
$aboutsubtitle.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Italic)

# Label above input
$abouttext                      = New-Object System.Windows.Forms.TextBox
$abouttext.Text                 = $text.About.abouttext
$abouttext.Size                 = New-Object System.Drawing.Size(255,165)
$abouttext.Left                 = ($GUI_Form_MainWindow_leftalign + 25)
$abouttext.Top                  = 150
$abouttext.ReadOnly             = $true
$abouttext.BackColor            = "White"
$abouttext.Multiline            = $true
$abouttext.TextAlign            = "Center"
$abouttext.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Regular)




$gotogithub                     = New-Object System.Windows.Forms.Button
$gotogithub.Size                = New-Object System.Drawing.Size (100,25)
$gotogithub.Left                = ($GUI_Form_MainWindow_leftalign)
$gotogithub.Top                 = $buttonalign
$gotogithub.Text                = $text.About.button_repo
$gotogithub.Add_Click( {start-process "https://github.com/teamcons/Skrivanek-Rocketlaunch"} )

$gotolicense                    = New-Object System.Windows.Forms.Button
$gotolicense.Size               = New-Object System.Drawing.Size (100,25)
$gotolicense.Left               = ($GUI_Form_MainWindow_leftalign + 105)
$gotolicense.Top                = $buttonalign
$gotolicense.Text               = $text.About.button_licence
$gotolicense.Add_Click( {start-process "https://www.gnu.org/licenses/gpl-3.0.html"})


$supportme                      = New-Object System.Windows.Forms.Button
$supportme.Size                 = New-Object System.Drawing.Size (100,25)
$supportme.Left                 = ($GUI_Form_MainWindow_leftalign + 210)
$supportme.Top                  = $buttonalign
$supportme.Text                 = $text.About.button_support
$supportme.Add_Click( {start-process "https://ko-fi.com/teamcons"})


$GUI_Tab_About.Controls.Add($abouttitle)
$GUI_Tab_About.Controls.Add($aboutsubtitle)
$GUI_Tab_About.Controls.Add($applogo)
$GUI_Tab_About.Controls.Add($abouttext)
$GUI_Tab_About.Controls.Add($gotogithub)
$GUI_Tab_About.Controls.Add($supportme)
$GUI_Tab_About.Controls.Add($gotolicense)


$GUI_Form_MainWindowTabControl.Controls.Add($GUI_Tab_About)

