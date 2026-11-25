$port = 8080
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  ContSeal Checker (v12) dang chay tai: http://localhost:$port" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan

$mimeTypes = @{
  ".html" = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".svg"  = "image/svg+xml"
}

while ($listener.IsListening) {
  $context = $listener.GetContext()
  $reqUrl = $context.Request.Url.LocalPath
  if ($reqUrl -eq "/api/log" -and $context.Request.HttpMethod -eq "POST") {
    $reader = New-Object System.IO.StreamReader($context.Request.InputStream, [System.Text.Encoding]::UTF8)
    $body = $reader.ReadToEnd()
    $reader.Close()
    Write-Host "[CLIENT LOG] $body" -ForegroundColor Yellow
    $context.Response.StatusCode = 200
    $context.Response.OutputStream.Close()
    continue
  }
  if ($reqUrl -eq "/api/archive" -and $context.Request.HttpMethod -eq "POST") {
    try {
      $reader = New-Object System.IO.StreamReader($context.Request.InputStream, [System.Text.Encoding]::UTF8)
      $body = $reader.ReadToEnd()
      $reader.Close()

      $stamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
      $parsed = $body | ConvertFrom-Json
      $rawVessel = if ($parsed.vesselName) { [string]$parsed.vesselName } else { "ChuyenTau" }
      $safeVessel = ($rawVessel -replace '\.(xlsx|xls|csv)$', '') -replace '[^A-Za-z0-9_-]', '_'
      if (-not $safeVessel) { $safeVessel = "ChuyenTau" }

      $archiveRoot = Join-Path $root "Archives"
      $sessionDir = Join-Path $archiveRoot "Tau_${safeVessel}_$stamp"
      $imgDir = Join-Path $sessionDir "Images"
      New-Item -ItemType Directory -Force -Path $imgDir | Out-Null

      $jsonPath = Join-Path $sessionDir "Archive_Tau_${safeVessel}_$stamp.json"
      [System.IO.File]::WriteAllText($jsonPath, $body, [System.Text.Encoding]::UTF8)

      if ($parsed.htmlReport) {
        $htmlPath = Join-Path $sessionDir "BaoCao_Tau_${safeVessel}_$stamp.html"
        [System.IO.File]::WriteAllText($htmlPath, [string]$parsed.htmlReport, [System.Text.Encoding]::UTF8)
      }

      $savedImgCount = 0
      if ($parsed.records) {
        foreach ($rec in $parsed.records) {
          $safeCont = ($rec.contNo -replace '[^A-Za-z0-9_-]', '_')
          $safeId = ($rec.id -replace '[^A-Za-z0-9_-]', '_')
          if ($rec.contImage -and $rec.contImage -match '^data:image/\w+;base64,(.+)$') {
            $b64 = $Matches[1]
            $imgBytes = [System.Convert]::FromBase64String($b64)
            $imgFile = Join-Path $imgDir "${safeCont}_Cont_${safeId}.jpg"
            [System.IO.File]::WriteAllBytes($imgFile, $imgBytes)
            $savedImgCount++
          }
          if ($rec.sealImage -and $rec.sealImage -match '^data:image/\w+;base64,(.+)$') {
            $b64 = $Matches[1]
            $imgBytes = [System.Convert]::FromBase64String($b64)
            $imgFile = Join-Path $imgDir "${safeCont}_Seal_${safeId}.jpg"
            [System.IO.File]::WriteAllBytes($imgFile, $imgBytes)
            $savedImgCount++
          }
        }
      }

      Write-Host "[ARCHIVE SAVED] $sessionDir ($savedImgCount images)" -ForegroundColor Green
      $respObj = @{ ok = $true; folder = $sessionDir; imageCount = $savedImgCount } | ConvertTo-Json -Compress
      $respBytes = [System.Text.Encoding]::UTF8.GetBytes($respObj)
      $context.Response.ContentType = "application/json; charset=utf-8"
      $context.Response.ContentLength64 = $respBytes.Length
      $context.Response.StatusCode = 200
      $context.Response.OutputStream.Write($respBytes, 0, $respBytes.Length)
    } catch {
      $errObj = @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress
      $errBytes = [System.Text.Encoding]::UTF8.GetBytes($errObj)
      $context.Response.ContentType = "application/json; charset=utf-8"
      $context.Response.ContentLength64 = $errBytes.Length
      $context.Response.StatusCode = 500
      $context.Response.OutputStream.Write($errBytes, 0, $errBytes.Length)
    }
    $context.Response.OutputStream.Close()
    continue
  }
  if ($reqUrl -eq "/") { $reqUrl = "/index.html" }
  $filePath = Join-Path $root ($reqUrl.TrimStart('/').Replace('/', '\'))

  $context.Response.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate")
  $context.Response.Headers.Add("Pragma", "no-cache")
  $context.Response.Headers.Add("Expires", "0")

  if (Test-Path $filePath -PathType Leaf) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    $contentType = $mimeTypes[$ext]
    if (-not $contentType) { $contentType = "application/octet-stream" }
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    $context.Response.ContentType = $contentType
    $context.Response.ContentLength64 = $bytes.Length
    $context.Response.StatusCode = 200
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } else {
    $context.Response.StatusCode = 404
  }
  $context.Response.OutputStream.Close()
}
