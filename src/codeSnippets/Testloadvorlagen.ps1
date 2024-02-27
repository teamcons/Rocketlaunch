


[string]$TEMPLATE               = "vorlagen.csv"
[string]$TEMPLATEDELIMITER               = ";"


if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
{ 
   $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
}
else
{ 
   $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
   if (!$ScriptPath){ $ScriptPath = "." } 
}



try {
    $detectedtemplate = (Import-Csv -Delimiter $TEMPLATEDELIMITER -Path (-join($ScriptPath,"\",$TEMPLATE))  -Header "Name","00","01","02","03","04","05","06","07","08","09")
    foreach ($row in $detectedtemplate)
    {
        $templates.Rows.Add($row."Name",$row."00",$row."01",$row."02",$row."03",$row."04",$row."05",$row."06",$row."07",$row."08",$row."09");
    }
}
catch {
    Write-Output "[ERROR] Cannot load templates, falling back to default"
    $templates.Rows.Add("Minimal","info","orig");
}
