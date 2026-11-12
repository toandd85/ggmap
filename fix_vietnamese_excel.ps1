param (
    [string]$InputCsv = ""
)

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)

function Fix-File([string]$path) {
    if (Test-Path $path) {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        # Check if already has BOM (EF BB BF)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            Write-Host "File da co UTF-8 BOM: $path" -ForegroundColor Cyan
            return
        }
        $text = [System.IO.File]::ReadAllText((Resolve-Path $path), [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText((Resolve-Path $path), $text, $utf8WithBom)
        
        # Tao them ban _excel.csv
        $excelCsv = [System.IO.Path]::ChangeExtension($path, "")
        if (-not $excelCsv.EndsWith("_excel.")) {
            $excelCsv = $excelCsv.TrimEnd('.') + "_excel.csv"
            [System.IO.File]::WriteAllText($excelCsv, $text, $utf8WithBom)
        }
        Write-Host "Da chuan hoa UTF-8 BOM: $path" -ForegroundColor Green
    }
}

if ($InputCsv -ne "") {
    Fix-File $InputCsv
} else {
    Write-Host "Dang chuan hoa toan bo file CSV trong thu muc va thu muc outputs..." -ForegroundColor Yellow
    Get-ChildItem -Path "." -Filter "*.csv" -File | ForEach-Object { Fix-File $_.FullName }
    if (Test-Path "outputs") {
        Get-ChildItem -Path "outputs" -Filter "*.csv" -File | ForEach-Object { Fix-File $_.FullName }
    }
    Write-Host "Hoan tat chuan hoa! Gio day mo truc tiep bang Excel se hien thi 100% tieng Viet co dau khong bao gio bi loi font." -ForegroundColor Green
}
