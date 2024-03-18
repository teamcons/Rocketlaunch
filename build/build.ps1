
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }

# Generate main exe
ps2exe `
-inputFile $ScriptPath\..\main.ps1 `
-iconFile $ScriptPath\..\assets\icon.ico `
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
-outputFile $ScriptPath\..\"Start Rocketlaunch.exe"


# Generate executable
ps2exe `
-inputFile $ScriptPath\install-rocketlaunch.ps1 `
-iconFile $ScriptPath\..\assets\icon.ico `
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
-outputFile $ScriptPath\..\"Install Rocketlaunch.exe"