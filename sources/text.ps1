﻿


        #=====================================
        #                TEXT                =
        #=====================================

<# 
All text strings
Maybe for localization, load from a CSV ?...
Powershell builtin tools arent super KISS to my taste...

#>



#================================================================
Write-Output "[START] Loading text"

[string]$script:LANG                            = $PSUICulture
switch -Wildcard ($LANG) {

   "*" {

        #================================
        ### MAIN UI
        [string]$global:text_projectname                    = "Projekt bereit zum Abflug!"
        [string]$global:text_doanalysis                     = "Analyse machen ? (Langsam)"
        [string]$global:text_opentrados                     = "Trados?"
        [string]$global:text_keepontop                      = "Über alle Fenster"

        [string]$global:text_splash_loadingapp              = "Ladung der Module"
        [string]$global:text_splash_loadingoutlook          = "Ladet Outlook..."
        [string]$global:text_splash_loading                 = "Ladet email: "

        #================================
        # Listview
        [string]$global:text_loadfilesfrom                  = 'Ausgangsdatei aus Quelle'

        # OUTLOOK
        [string]$global:text_label_from_Outlook             = "Ausgangsdatei unter alle E-Mails mit Anhang seit dem Vortag, 17:30 Uhr"
        [string]$global:text_from_Outlook                   = "Outlook Emails"
        [string]$global:text_columns_Subject                = 'Betreff'
        [string]$global:text_columns_Sendername             = 'Von'
        [string]$global:text_columns_Attachments            = 'Dateien'
        [string]$global:text_columns_time                   = 'Ankunft'
        [string]$global:text_sourcefiles_refresh            = 'Erneut laden'

        # DOWNLOADS
        [string]$global:text_label_from_Downloads           = "Alle Dateien, die seit heute heruntergeladen wurden"
        [string]$global:text_from_Downloads                 = "Downloads"
        [string]$global:text_columns_File                   = 'Datei'
        [string]$global:text_columns_Directory              = 'Ordner'
        [string]$global:text_columns_LastWrite              = 'Letzte Änderung'

        # DRAG N DROP
        [string]$global:text_DragNDrop                      = "Ziehen & Ablegen"
        [string]$global:text_columns_Path                   = 'Vollständiger Pfad'
        [string]$global:text_label_DragNDrop                = "Alle in der Rasteransicht abgelegten Dateien"


        # NONE WITH LEFT BEEF
        [string]$global:text_label_nofilesource             = "Keine Ausgangsdatei einbeziehen"
        [string]$global:text_nofilesource                   = "Keine Ausgangsdatei"


        [string]$global:text_nooutlook                   = "Kein Outlook Ordner"



        #================================
        # Datagridview, templates
        [string]$global:text_usewhichtemplate               = 'Welche Projektvorlage verwenden?' #(Tipp: Namen eines Ordners durch Doppelklicken ändern)'
        [string]$global:text_savetemplatechanges            = "Änderungen speichern"
        [string]$global:text_template_name                  = "Vorlage"
        [string]$global:text_help                           = "Erweitert"
        [string]$global:text_OK                             = "Los!"
        [string]$global:text_Cancel                         = "Nö"
        [string]$global:text_NotifyText                     = "Projekt ist bereit !"


        [string]$global:text_csv_file                     = "Datei"
        [string]$global:text_csv_wordcount                    = "Wortzahl"
        [string]$global:text_csv_proofreadtime             = "Überprüfungszeit"
        [string]$global:text_csv_total                     = "TOTAL"

        #================================
        ### Settings tab
        [string]$global:text_settingstag                    = "Erweiterte Einstellungen"
        [string]$global:text_settingsnota                   = "Sie werden nicht gespeichert!"

        # OPTIONS
        [string]$global:text_settings_ExplorerQuickAccess   = "Eine Schnellzugriffsverknüpfung im Explorer erstellen ?"
        [string]$global:text_settings_OutlookFolder         = "Einen Ordner in Outlook erstellen?"
        [string]$global:text_settings_Countwords            = "Wörter zählen ?"
        [string]$global:text_settings_OpenExplorer          = "Neu erstellten Ordner nach Fertigstellung öffnen ?"
        [string]$global:text_settings_createarchivefolder         = "ArchivenOrdner in Delivered Erstellen wenn nötig ?"
        [string]$global:text_settings_Notify                = "Eine Benachrichtigung senden, wenn fertig ?"
        [string]$global:text_settings_CloseAfter            = "Programm nach der Erstellung verlassen?"

        [string]$script:text_label_select_lang              = "Sprache ändern (TODO)"
        [string]$script:text_label_select_theme             = "Thema ändern (TODO)"
        [string]$script:text_lang_german                    = "Deutsch"
        [string]$script:text_lang_french                    = "Französisch"
        [string]$script:text_lang_spanish                   = "Spanish"

        # HELP
        [string]$global:text_settings_helptitle             = "Hilfe"
        [string]$global:text_settings_getthedoc             = "Dokumentation Öffnen"
        [string]$global:text_settings_askme                 = "Mich fragen"

        [string]$global:text_settings_close                 = "Ok!"

        #================================
        # ABOUT TAB
        [string]$global:text_abouttab                       = "Stella!" 
        [string]$global:text_aboutsubtitle                  = "Neue Projekte Starten, aber sehr, sehr schnell!"
        [string]$global:text_abouttext                      = "Mit Liebe gemacht von Stella,
für ihre Arbeit bei Skrivanek GmbH
        
Ich hoffe ihr findet es nützlich!
Ich bin kein Entwickler, ich habe Wirtschaft studiert, ich habe keine Ahnung von diesen Geek-Sachen.
        
Version 2.0.irgendwas
2024 Stella Ménier, unter GNU GPL v3"
        [string]$global:text_about_button_repo              = "Project repo"
        [string]$global:text_about_button_licence           = "Licence"
        [string]$global:text_about_button_support           = "Support me!"

    } # DE
} # End of Big Switch







<# 
= "Start new projects, but very very quickly !"

[string]$global:text_abouttext             = "Made with love by Stella,
for her work at Skrivanek GmbH

I hope you find it useful !
I am no developer, i studied economics, ive got no clue of those geek things.

Version 2.0.somethingsomething
2024 Stella Ménier, under GNU GPL v3" #>