
#========================================
# Get all important variables in place 


    #Write-Output "[STARTUP] Getting all variables in place"
    [string]$script:APPNAME                         = "-Rocketlaunch!"
    [string]$script:ROOTSTRUCTURE                   = "M:\9_JOBS_XTRF\"
    [string]$script:YEAR                            = get-date –f yyyy
    [regex]$script:CODEPATTERN                      = -join($YEAR,"-[0-9]")
    
    [string]$script:LOAD_TEMPLATES_FROM             = $ScriptPath
    [string]$script:TEMPLATE                        = "vorlagen.csv"
    [string]$script:TEMPLATEDELIMITER               = ';'
    
    [string]$script:default_filesfrom               = $text_from_Outlook
    #[string]$global:default_fromdisk                = '$env:USERPROFILE\Downloads'
    [bool]$script:default_opentrados                = $true
    [bool]$script:default_expandarchives            = $true
    [bool]$script:default_createshortcut            = $true
    [bool]$script:default_createoutlookfolder       = $true
    [bool]$script:default_movesourcemail            = $true
    [bool]$script:default_openexplorer              = $true
    [bool]$script:default_notifywhenfinished        = $true

