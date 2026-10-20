# Export docx/xlsx files to PDF with local Microsoft Office (COM), so the caller
# can render the PDF pages to PNG with PyMuPDF for visual inspection.
# Usage: powershell -ExecutionPolicy Bypass -File render_check.ps1 -InDir tests\out -OutDir tests\preview -Filter *tender*

param(
    [string]$InDir = "tests\out",
    [string]$OutDir = "tests\preview",
    [string]$Filter = "*",
    [int]$MaxFiles = 20
)

$ErrorActionPreference = "Stop"
$InDir  = (Resolve-Path $InDir).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

$files = Get-ChildItem -Path $InDir | Where-Object { ($_.Extension -eq ".docx" -or $_.Extension -eq ".xlsx") -and $_.Name -like $Filter } | Select-Object -First $MaxFiles
if (-not $files) { Write-Host "no matching file"; exit 1 }

$word = $null
$excel = $null
try {
    foreach ($f in $files) {
        $pdf = [string](Join-Path $OutDir ($f.BaseName + ".pdf"))
        try {
            if ($f.Extension -eq ".xlsx") {
                if (-not $excel) {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                }
                $wb = $excel.Workbooks.Open($f.FullName, 0, $true)
                $wb.ExportAsFixedFormat(0, $pdf)
                $wb.Close($false)
                Write-Host ("OK  {0}" -f $f.Name)
            } else {
                if (-not $word) {
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0
                }
                $doc = $word.Documents.Open($f.FullName, $false, $true)
                $doc.SaveAs2($pdf, 17)   # 17 = wdFormatPDF
                $pages = $doc.ComputeStatistics(2)
                $doc.Close($false)
                Write-Host ("OK  {0}  pages={1}" -f $f.Name, $pages)
            }
        } catch {
            Write-Host ("ERR {0}  {1}" -f $f.Name, $_.Exception.Message)
        }
    }
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
