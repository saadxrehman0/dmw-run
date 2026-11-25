# Export the day's DMW decks to PDF with PowerPoint (read-only, no window). 1 Nov 2026.
# Usage: powershell -ExecutionPolicy Bypass -File export_pdf.ps1 -Folder "<day folder>"
param([Parameter(Mandatory=$true)][string]$Folder)
$ErrorActionPreference = 'Stop'
$names = @('ANB Daily Market Watch - EN', 'ANB Daily Market Watch - AR')
$app = New-Object -ComObject PowerPoint.Application
foreach ($n in $names) {
  $src = Join-Path $Folder ($n + '.pptx')
  $dst = Join-Path $Folder ($n + '.pdf')
  if (-not (Test-Path -LiteralPath $src)) { Write-Output "MISSING $src"; continue }
  # Open(FileName, ReadOnly, Untitled, WithWindow)
  $p = $app.Presentations.Open($src, -1, 0, 0)
  try {
    $slides = $p.Slides.Count
    # SaveAs(FileName, ppSaveAsPDF = 32)
    $p.SaveAs($dst, 32)
    $fi = Get-Item -LiteralPath $dst
    Write-Output ("OK {0} | slides {1} | pdf {2:N0} bytes | saved {3:HH:mm:ss}" -f $n, $slides, $fi.Length, $fi.LastWriteTime)
  } finally {
    $p.Close()
  }
}
if ($app.Presentations.Count -eq 0) { $app.Quit() }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
