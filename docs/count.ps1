




$word = New-Object -ComObject Word.Application 
$word.Visible = $false 
$file = $word.Documents.Open("M:\4_BE\06_General information\Stella\Skrivanek-Rocketlaunch\docs\Manual - Rocketlaunch.docx")

$lines = $file.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticLines)
$words = $file.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticWords) 
$chars = $file.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticCharacters)

Write-Host "Lines: $lines" 
Write-Host "Words: $words" 
Write-Host "Characters: $chars"

$file.Close() 
$word.Quit()
