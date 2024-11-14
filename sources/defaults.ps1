

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


    # CSV shenanigans
    [string]$script:default_csv_analysis            = "\Rocketlaunch-analysis.csv"
    [string]$script:TEMPLATEDELIMITER               = ';'
    [string]$script:SEP                             = ';'
    [int]$script:WORDS_PER_HOUR                     = 1800
    [int]$script:DECIMALS                           = 2


    # Exposed in main UI
    [string]$script:default_filesfrom               = $text.Sourceview.from_Outlook
    [bool]$script:default_ontop                     = $true
    [bool]$script:default_opentrados                = $true
    [bool]$script:default_savetemplatechanges       = $false

    # Exposed in settings
    [bool]$script:default_createshortcut            = $true
    [bool]$script:default_createoutlookfolder       = $true
    [bool]$script:default_movesourcemail            = $true
    [bool]$script:default_createarchivefolder       = $true # True whenever can figure out detect
    [bool]$script:default_openexplorer              = $true
    [bool]$script:default_notifywhenfinished        = $true
    [bool]$script:default_closeafter                = $true
    [bool]$script:default_countwords                = $false
    [bool]$script:default_restart                   = $false


    # Orange, Lightblue, Brushed Metal+LightGray - Aqua ?
    [string]$script:THEME                           = "Modern Color"

    # Do we still use it ???
    #[string]$script:supported_filetypes               = $text_from_Outlook


