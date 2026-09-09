<#
    Собирает содержимое APK из app/index.html.

    В браузере страница тянет шрифты и SheetJS из сети — в приложении так
    нельзя, оно должно работать офлайн и без разрешения на интернет. Скрипт
    складывает эти файлы в assets и переписывает ссылки на локальные.

    Запуск:  powershell -ExecutionPolicy Bypass -File android\sync-assets.ps1
    Ключ -SkipDownload пересобирает только index.html, не трогая загрузки.
#>
param([switch]$SkipDownload)

$ErrorActionPreference = "Stop"

$root    = Split-Path -Parent $PSScriptRoot
$src     = Join-Path $root "app\index.html"
$assets  = Join-Path $PSScriptRoot "app\src\main\assets"
$fontDir = Join-Path $assets "fonts"
$vendor  = Join-Path $assets "vendor"

$XLSX_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
$FONT_URL = "https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@500;700&family=Onest:wght@400;500;600;700&family=Unbounded:wght@700&display=swap"
$UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

# Русскому тексту хватает подмножеств cyrillic и latin; остальные не тянем.
$KEEP_SUBSETS = @("cyrillic", "latin")

New-Item -ItemType Directory -Force -Path $assets, $fontDir, $vendor | Out-Null

function Save-Utf8($path, $text) {
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

# ---------------------------------------------------------------- SheetJS
$xlsxFile = Join-Path $vendor "xlsx.full.min.js"
if (-not $SkipDownload -or -not (Test-Path $xlsxFile)) {
    Write-Host "Скачиваю SheetJS..."
    Invoke-WebRequest -Uri $XLSX_URL -OutFile $xlsxFile -UseBasicParsing
}

# ----------------------------------------------------------------- шрифты
$fontCssFile = Join-Path $fontDir "fonts.css"
if (-not $SkipDownload -or -not (Test-Path $fontCssFile)) {
    Write-Host "Скачиваю шрифты..."
    $css = (Invoke-WebRequest -Uri $FONT_URL -UserAgent $UA -UseBasicParsing).Content

    # Google отдаёт блоки, помеченные комментарием с названием подмножества.
    $blocks = [regex]::Matches($css, '/\*\s*([a-z0-9\-]+)\s*\*/\s*(@font-face\s*\{[^}]*\})')
    if ($blocks.Count -eq 0) { throw "Не удалось разобрать CSS шрифтов" }

    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($b in $blocks) {
        if ($KEEP_SUBSETS -notcontains $b.Groups[1].Value) { continue }
        $face = $b.Groups[2].Value

        foreach ($u in [regex]::Matches($face, 'url\((https://fonts\.gstatic\.com/[^)]+\.woff2)\)')) {
            $url  = $u.Groups[1].Value
            $name = ($url -split '/')[-1]
            $dst  = Join-Path $fontDir $name
            if (-not (Test-Path $dst)) {
                Invoke-WebRequest -Uri $url -OutFile $dst -UseBasicParsing
            }
            $face = $face.Replace($url, $name)
        }
        $kept.Add($face)
    }

    Save-Utf8 $fontCssFile ("/* Собрано android\sync-assets.ps1 из Google Fonts */`r`n" + ($kept -join "`r`n"))
    Write-Host ("  шрифтовых начертаний: " + $kept.Count)
}

# -------------------------------------------------------------- index.html
$html = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)

$title = [regex]::Match($html, '<title>(.*?)</title>').Groups[1].Value
if (-not $title) { $title = "Мультисекундомер" }

$body = $html
$body = [regex]::Replace($body, '<title>.*?</title>\s*', '')
$body = [regex]::Replace($body, '<link rel="preconnect"[^>]*>\s*', '')
$body = [regex]::Replace($body, '<link rel="stylesheet" href="https://fonts\.googleapis\.com[^"]*">\s*', '')
$body = [regex]::Replace($body, '<script src="https://cdnjs\.cloudflare\.com[^"]*"></script>\s*', '')

if ($body -match 'https://(fonts|cdnjs)\.') { throw "В странице осталась ссылка в сеть — проверьте шаблоны замены" }

$doc = @"
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>$title</title>
<link rel="stylesheet" href="fonts/fonts.css">
<style>
  :root { color-scheme: light dark; }
  body { margin: 0; }
  img { max-width: 100%; }
  [hidden] { display: none !important; }
</style>
<script src="vendor/xlsx.full.min.js"></script>
</head>
<body>
$body
</body>
</html>
"@

Save-Utf8 (Join-Path $assets "index.html") $doc
Write-Host "Готово: $assets"
