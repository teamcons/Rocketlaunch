
#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$global:APPNAME                         = "-Rocketlaunch!"
[string]$global:ROOTSTRUCTURE                   = "M:\9_JOBS_XTRF\"
[string]$global:YEAR                            = get-date –f yyyy
[regex]$global:CODEPATTERN                      = -join($YEAR,"-[0-9]")

[string]$global:LOAD_TEMPLATES_FROM             = $ScriptPath
[string]$global:TEMPLATE                        = "vorlagen.csv"
[string]$global:TEMPLATEDELIMITER               = ";"

[string]$global:default_filesfrom               = $text_from_Outlook
[string]$global:default_fromdisk                = "$env:USERPROFILE\Downloads\"
[bool]$global:default_opentrados                = $true
[bool]$global:default_expandarchives            = $true
[bool]$global:default_createshortcut            = $true
[bool]$global:default_createoutlookfolder       = $true
[bool]$global:default_movesourcemail            = $true
[bool]$global:default_openexplorer              = $true
[bool]$global:default_notifywhenfinished        = $true


