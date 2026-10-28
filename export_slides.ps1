$ppt = New-Object -ComObject PowerPoint.Application
$pptxPath = Join-Path (Get-Location) "Student_Record_Management_CA2.pptx"
$pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$outDir = Join-Path (Get-Location) "screenshots\slides"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $outPath = Join-Path $outDir ("slide_" + $i + ".png")
    $pres.Slides.Item($i).Export($outPath, "PNG")
    Write-Host "Exported Slide $i to $outPath"
}

$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "All slides exported successfully."
