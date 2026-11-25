# Export DMW v2 decks to PDF with PowerPoint and record PowerPoint's own table layout.
# Usage: powershell -ExecutionPolicy Bypass -File export_measure.ps1 -Folder "<folder>" [-Recurse]
# For every "ANB Daily Market Watch - EN/AR.pptx" under the folder: saves the PDF beside it and adds
# each table row on slides 1-3 (top, height, row height) to "<folder>\ppt_layout.csv", so the run can
# check the 10.60 in rule on PowerPoint's real layout (fonts, wraps) instead of an estimate.
param([Parameter(Mandatory=$true)][string]$Folder, [switch]$Recurse)
$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
$rows = New-Object System.Collections.Generic.List[object]
$files = Get-ChildItem -LiteralPath $Folder -Filter 'ANB Daily Market Watch - *.pptx' -Recurse:$Recurse
foreach ($f in $files) {
    $pres = $pp.Presentations.Open($f.FullName, $true, $false, $true)   # read-only, with window so layout runs
    try {
        foreach ($s in $pres.Slides) {
            if ($s.SlideIndex -gt 3) { break }
            foreach ($sh in $s.Shapes) {
                if (-not $sh.HasTable) { continue }
                $t = $sh.Table
                for ($r = 1; $r -le $t.Rows.Count; $r++) {
                    $rows.Add([pscustomobject]@{
                        file = $f.FullName.Substring($Folder.Length).TrimStart('\'); slide = $s.SlideIndex
                        shape = $sh.Name; top_in = [math]::Round($sh.Top / 72, 4)
                        height_in = [math]::Round($sh.Height / 72, 4); bottom_in = [math]::Round(($sh.Top + $sh.Height) / 72, 4)
                        row = $r; row_h_in = [math]::Round($t.Rows.Item($r).Height / 72, 4)
                    })
                }
            }
        }
        $pdf = [IO.Path]::ChangeExtension($f.FullName, '.pdf')
        $pres.SaveAs($pdf, 32)
        Write-Output "OK $($f.FullName)"
    } finally { $pres.Close() }
}
$rows | Export-Csv -LiteralPath (Join-Path $Folder 'ppt_layout.csv') -NoTypeInformation -Encoding UTF8
if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
Write-Output "DONE $($files.Count) decks"
