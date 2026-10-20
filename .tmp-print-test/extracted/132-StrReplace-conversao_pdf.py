OLD:
try {
    $brutos = Get-Content -LiteralPath $Mapa -Encoding UTF8 -Raw | ConvertFrom-Json

NEW:
$script:excelAntes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
try {
    $brutos = Get-Content -LiteralPath $Mapa -Encoding UTF8 -Raw | ConvertFrom-Json