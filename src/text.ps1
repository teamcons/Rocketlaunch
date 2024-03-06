

Write-Output "[START] Loading text"


    ### MAIN UI
    [string]$global:text_projectname               = "Projekt bereit zum Abflug!"
    [string]$global:text_doanalysis                = "Analyse machen ? (Langsam)"
    [string]$global:text_opentrados                = "Trados?"
    [string]$global:text_keepontop                 = "Über alle Fenster"

    # Listview
    [string]$global:text_loadfilesfrom             = 'Ausgangsdatei aus Quelle'
    [string]$global:text_columns_Subject           = 'Betreff'
    [string]$global:text_columns_Sendername        = 'Von'
    [string]$global:text_columns_Attachments       = 'Dateien'
    [string]$global:text_columns_time              = 'Ankunft'
    [string]$global:text_from_Outlook              = "In Outlook"
    [string]$global:text_from_Downloads            = "(TODO) In Downloads"
    [string]$global:text_DragNDrop                 = "(TODO) Drag&Drop"
    [string]$global:text_nofilesource              = "(TODO) Keine Ausgangsdatei"

    
    # Datagridview, templates
    [string]$global:text_usewhichtemplate          = 'Welche Projektvorlage soll verwendet werden?'
    [string]$global:text_loadtemplate              = "Laden..."
    [string]$global:text_help                      = "Mehr"
    [string]$global:text_OK                        = "Los!"
    [string]$global:text_Cancel                    = "Nö"


    ### Settings tab
    [string]$global:text_settingstag                   = "Erweiterte Einstellungen" 
    [string]$global:text_settings_ExplorerQuickAccess  = "Create a quick access shortcut in Explorer ?"
    [string]$global:text_settings_OutlookFolder        = "Create a folder in Outlook ?"
    [string]$global:text_settings_OpenExplorer         = "Open newly created folder when finished ?"
    [string]$global:text_settings_Notify               = "Send a notification when done ?"
    [string]$global:text_settings_helptitle            = "Help"
    [string]$global:text_settings_getthedoc            = "Download latest manual version"
    [string]$global:text_settings_askme                = "Ask me"
    

    ### ABOUT TAB
    [string]$global:text_abouttab              = "Stella!" 
    [string]$global:text_aboutsubtitle         = "Start new projects, but very very quickly !"
    [string]$global:text_abouttext             = "Made with love by Stella, for Skrivanek GmbH

I hope you find it useful !
I am no developer, i studied economics, ive got no clue of those geek things.

Version 2.0.somethingsomething
2024 Stella Ménier, under GNU GPL v3"
    [string]$global:text_about_button_repo     = "Project repo"
    [string]$global:text_about_button_licence  = "Licence"
    [string]$global:text_about_button_support  = "Support me!"
