

        #=========================================
        #                DEFAULTS                =
        #=========================================

<# 
All the default environment values
Some are obsolete, some would maybe be better in limited scope
Some arent really supposed to be changed
#>

#========================================
# Get all important variables in place 

Write-Output "[START] Loading defaults"

    #Write-Output "[STARTUP] Getting all variables in place"
    [string]$script:APPNAME                         = "-Rocketlaunch!"
    [string]$script:ROOTSTRUCTURE                   = "M:\9_JOBS_XTRF\"
    [string]$script:YEAR                            = get-date –f yyyy
    [regex]$script:CODEPATTERN                      = -join($YEAR,"-[0-9]")

    [regex]$script:accepted_attachments             = ".(pdf|doc|docx|xls|xlsx|ppt|pptx|xml|idml|csv|txt|zip|sdlppx)"
    [regex]$script:unsupported                      = ".(jpg|png|gif|webp|jpeg)"


        [string]$script:SEP                      = ";"

    # Exposed in main UI
    [string]$script:default_filesfrom               = $text.Sourceview.from_Outlook
    [bool]$script:default_ontop                     = $settings.UI.KeepOnTop
    [bool]$script:default_opentrados                = $settings.Preferences.opentrados
    [bool]$script:default_savetemplatechanges       = $settings.Preferences.savetemplatechanges

    # Exposed in settings
    [bool]$script:default_createshortcut            = $settings.Preferences.createshortcut
    [bool]$script:default_createoutlookfolder       = $settings.Preferences.createoutlookfolder
    [bool]$script:default_movesourcemail            = $settings.Preferences.movesourcemail
    [bool]$script:default_createarchivefolder       = $settings.Preferences.createarchivefolder
    [bool]$script:default_openexplorer              = $settings.Preferences.openexplorer
    [bool]$script:default_notifywhenfinished        = $settings.Preferences.notifywhenfinished
    [bool]$script:default_closeafter                = $settings.Preferences.keepopen


    # Orange, Lightblue, Brushed Metal+LightGray - Aqua ?
    [string]$script:THEME                           = $settings.UI.Theme
