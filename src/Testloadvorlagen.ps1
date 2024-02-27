


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


$detectedtemplate = (Import-Csv -Delimiter $TEMPLATEDELIMITER -Path (-join($ScriptPath,"\",$TEMPLATE))  -Header "Name","00","01","02","03","04","05","06","07","08","09" | Format-Table)


echo $detectedtemplate.Rows
