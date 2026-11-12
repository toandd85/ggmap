param (
    [string]$InputCsv = "results.csv",
    [string]$OutputCsv = ""
)

if (-not (Test-Path $InputCsv)) {
    Write-Host "Khong tim thay file $InputCsv" -ForegroundColor Red
    exit 1
}

if ($OutputCsv -eq "") {
    $OutputCsv = [System.IO.Path]::GetFileNameWithoutExtension($InputCsv) + "_excel.csv"
}

$text = [System.IO.File]::ReadAllText((Resolve-Path $InputCsv), [System.Text.Encoding]::UTF8)
$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText((Join-Path (Get-Location) $OutputCsv), $text, $utf8WithBom)

Write-Host "Da chuyen doi thanh cong: $OutputCsv (Chuan UTF-8 BOM, mo truc tiep bang Excel se hien thi dung 100% tieng Viet co dau)" -ForegroundColor Green
