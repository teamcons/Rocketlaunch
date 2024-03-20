
#========================================
# Get all important variables in place 

Write-Output "[START] Loading defaults"


    #Write-Output "[STARTUP] Getting all variables in place"
    [string]$script:APPNAME                         = "-Rocketlaunch!"

    [string]$script:ROOTSTRUCTURE                   = "M:\9_JOBS_XTRF\"
    [string]$script:YEAR                            = get-date –f yyyy
    [regex]$script:CODEPATTERN                      = -join($YEAR,"-[0-9]")

    [string]$script:TEMPLATEDELIMITER               = ';'




    
    [string]$script:supported_filetypes               = $text_from_Outlook

    [string]$script:default_filesfrom               = $text_from_Outlook
    #[string]$global:default_fromdisk                = '$env:USERPROFILE\Downloads'
    [bool]$script:default_ontop                = $false
    [bool]$script:default_opentrados                = $true
    [bool]$script:default_savetemplatechanges       = $true

    [bool]$script:default_createshortcut            = $true
    [bool]$script:default_createoutlookfolder       = $true
    [bool]$script:default_movesourcemail            = $true
    [bool]$script:default_openexplorer              = $true
    [bool]$script:default_notifywhenfinished        = $true
    [bool]$script:default_closeafter                = $true

    [bool]$script:default_countwords                = $false
    [bool]$script:default_restart                   = $false