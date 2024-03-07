
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }


ps2exe `
-inputFile $ScriptPath\launcher.ps1 `
-iconFile $ScriptPath\..\assets\Rocketlaunch-Icon.ico `
-noConsole `
-noOutput `
-exitOnCancel `
-title "-Rocketlaunch!" `
-description "Start new projects, but very very quickly !" `
-company "By Stella, for Skrivanek GmbH" `
-product "Rocketlaunch!" `
-copyright "GNU GPL-3.0 Stella <stella.menier@gmx.de>" `
-version 2.0 `
-Verbose `
-outputFile $ScriptPath\..\Start-Rocketlaunch.exe