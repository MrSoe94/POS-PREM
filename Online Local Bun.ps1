param(
    [int]$Port = 3011,
    [switch]$NoTunnel,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'

function Pause-IfNeeded {
    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Tekan Enter untuk menutup"
    }
}

function Get-PreferredLanIPv4 {
    # Prefer adapter yang Up + punya gateway (biasanya WiFi/Ethernet utama).
    try {
        $cfg = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
            Where-Object {
                $_.NetAdapter.Status -eq 'Up' -and
                $_.IPv4DefaultGateway -ne $null -and
                $_.IPv4Address -ne $null
            } |
            Select-Object -First 1
        if ($cfg -and $cfg.IPv4Address) {
            $ip = @($cfg.IPv4Address)[0].IPAddress
            if ($ip -and $ip -notmatch '^(127\.|169\.254\.)') { return $ip }
        }
    } catch {}

    try {
        $ip = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object {
                $_.IPAddress -notmatch '^(127\.|169\.254\.)' -and
                $_.PrefixOrigin -ne 'WellKnown'
            } |
            Sort-Object InterfaceMetric |
            Select-Object -ExpandProperty IPAddress -First 1
        if ($ip) { return $ip }
    } catch {}

    try {
        $ip = [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
            Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.ToString() -notmatch '^(127\.|169\.254\.)' } |
            Select-Object -ExpandProperty IPAddressToString -First 1
        if ($ip) { return $ip }
    } catch {}

    return $null
}

function Open-AccessQrPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LanUrl,
        [string]$TunnelUrl
    )

    $lanEncoded = [uri]::EscapeDataString($LanUrl)
    $lanQr = "https://api.qrserver.com/v1/create-qr-code/?size=360x360&data=$lanEncoded"

    $tunnelBlock = ''
    if ($TunnelUrl) {
        $tunnelEncoded = [uri]::EscapeDataString($TunnelUrl)
        $tunnelQr = "https://api.qrserver.com/v1/create-qr-code/?size=360x360&data=$tunnelEncoded"
        $tunnelEsc = [System.Net.WebUtility]::HtmlEncode($TunnelUrl)
        $tunnelBlock = @"
      <section class="card">
        <h2>Online (Tunnel)</h2>
        <p>Akses dari internet / luar jaringan</p>
        <img src="$tunnelQr" alt="QR Tunnel" width="360" height="360" />
        <a href="$tunnelEsc">$tunnelEsc</a>
      </section>
"@
        Write-Host "Link tunnel: $TunnelUrl" -ForegroundColor Green
    }

    $lanEsc = [System.Net.WebUtility]::HtmlEncode($LanUrl)
    Write-Host "Link lokal (LAN): $LanUrl" -ForegroundColor Green
    Write-Host "Membuka QR Code (LAN + Tunnel) di browser..." -ForegroundColor Cyan

    $html = @"
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>POS Access QR</title>
  <style>
    :root { color-scheme: light; }
    body {
      margin: 0; min-height: 100vh; font-family: Segoe UI, Tahoma, sans-serif;
      background: #f3f6fb; color: #10233f;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      gap: 20px; padding: 24px;
    }
    h1 { margin: 0; font-size: 1.35rem; }
    .hint { margin: 0; color: #4b5d78; text-align: center; max-width: 720px; }
    .row { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; }
    .card {
      background: #fff; border: 1px solid #d7e0ee; border-radius: 14px;
      padding: 18px 18px 14px; width: 400px; text-align: center;
      box-shadow: 0 8px 24px rgba(16, 35, 63, 0.06);
    }
    .card h2 { margin: 0 0 4px; font-size: 1.05rem; }
    .card p { margin: 0 0 12px; color: #4b5d78; font-size: 0.92rem; }
    .card img { display: block; margin: 0 auto 12px; border-radius: 8px; }
    .card a { word-break: break-all; color: #0b5fff; font-size: 0.88rem; }
  </style>
</head>
<body>
  <h1>Scan QR untuk buka aplikasi</h1>
  <p class="hint">LAN = HP/tablet di WiFi yang sama. Tunnel = akses online via Cloudflare.</p>
  <div class="row">
    <section class="card">
      <h2>Lokal (LAN / IP)</h2>
      <p>Jaringan yang sama — lebih ringan untuk cetak &amp; kasir</p>
      <img src="$lanQr" alt="QR LAN" width="360" height="360" />
      <a href="$lanEsc">$lanEsc</a>
    </section>
$tunnelBlock
  </div>
</body>
</html>
"@

    $htmlPath = Join-Path $env:TEMP ("pos-access-qr-" + [Guid]::NewGuid().ToString("N") + ".html")
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
    Start-Process $htmlPath | Out-Null
}

function Get-CloudflaredVersion {
    try {
        $raw = & cloudflared --version 2>$null
        if ($raw -match 'cloudflared version\s+([0-9]+(?:\.[0-9]+){1,3})') {
            return [version]$Matches[1]
        }
    } catch {}
    return $null
}

function Ensure-BunInstalled {
    if (Get-Command bun -ErrorAction SilentlyContinue) {
        return
    }

    Write-Host "Bun belum terinstall. Menginstall Bun..." -ForegroundColor Yellow
    powershell -c "irm bun.sh/install.ps1|iex"

    $bunBin = Join-Path $env:USERPROFILE '.bun\bin'
    if (Test-Path -LiteralPath $bunBin) {
        $env:PATH = "$bunBin;$env:PATH"
    }

    if (-not (Get-Command bun -ErrorAction SilentlyContinue)) {
        throw "Instalasi Bun selesai tetapi 'bun' belum ditemukan di PATH. Tutup dan buka ulang PowerShell, lalu jalankan script ini lagi."
    }

    $version = & bun --version 2>$null
    Write-Host "Bun berhasil terinstall (versi $version)." -ForegroundColor Green
}

try {
    Write-Host "Memulai Aplikasi (Bun)..." -ForegroundColor Cyan
    Write-Host "Lokasi project: $PSScriptRoot" -ForegroundColor DarkGray

    Ensure-BunInstalled

    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'server.js'))) {
        throw "server.js tidak ditemukan di folder project. Jalankan script dari folder aplikasi yang benar."
    }

    # Jalankan server dengan Bun di jendela PowerShell terpisah
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", "bun server.js") `
        -WorkingDirectory $PSScriptRoot |
        Out-Null

    Write-Host "Menunggu server aktif..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    $lanIp = Get-PreferredLanIPv4
    $lanUrl = if ($lanIp) { "http://${lanIp}:$Port" } else { "http://localhost:$Port" }
    if (-not $lanIp) {
        Write-Host "[WARN] IP LAN tidak terdeteksi. QR lokal memakai localhost." -ForegroundColor Yellow
    }

    if ($NoTunnel) {
        Write-Host "Mode lokal aktif. Buka: $lanUrl" -ForegroundColor Green
        Open-AccessQrPage -LanUrl $lanUrl
        Pause-IfNeeded
        return
    }

    if (-not (Get-Command cloudflared -ErrorAction SilentlyContinue)) {
        throw "cloudflared tidak ditemukan di PATH. Install cloudflared atau jalankan dengan -NoTunnel."
    }

    $cfVersion = Get-CloudflaredVersion
    if ($cfVersion) {
        Write-Host "Versi cloudflared: $cfVersion" -ForegroundColor DarkGray
        if ($cfVersion -lt [version]'2024.8.0') {
            Write-Host "[WARN] cloudflared cukup lama. Jika quick tunnel gagal, update ke versi terbaru." -ForegroundColor Yellow
        }
    }

    Write-Host "Memulai Cloudflare Tunnel..." -ForegroundColor Cyan
    $logGuid = [Guid]::NewGuid().ToString("N")
    $stdoutLog = Join-Path $env:TEMP ("cloudflared-" + $logGuid + "-out.log")
    $stderrLog = Join-Path $env:TEMP ("cloudflared-" + $logGuid + "-err.log")
    $process = Start-Process -FilePath "cloudflared" `
        -ArgumentList @("tunnel", "--url", "localhost:$Port") `
        -NoNewWindow `
        -RedirectStandardOutput $stdoutLog `
        -RedirectStandardError $stderrLog `
        -PassThru

    $outLineCount = 0
    $errLineCount = 0
    $qrOpened = $false
    $regex = 'https://[a-z0-9\.-]+\.trycloudflare\.com'

    $processNewLines = {
        param(
            [string]$Path,
            [ref]$Cursor
        )
        if (-not (Test-Path -LiteralPath $Path)) { return @() }
        $all = Get-Content -LiteralPath $Path
        if ($all.Count -le $Cursor.Value) { return @() }
        $start = $Cursor.Value
        $Cursor.Value = $all.Count
        return $all[$start..($all.Count - 1)]
    }

    while (-not $process.HasExited) {
        $newOut = & $processNewLines -Path $stdoutLog -Cursor ([ref]$outLineCount)
        $newErr = & $processNewLines -Path $stderrLog -Cursor ([ref]$errLineCount)
        $combined = @($newOut + $newErr)

        foreach ($line in $combined) {
            if ($line -ne $null -and $line -ne '') {
                Write-Host $line
                if (-not $qrOpened -and ($line -match $regex)) {
                    $qrOpened = $true
                    Open-AccessQrPage -LanUrl $lanUrl -TunnelUrl $Matches[0]
                }
            }
        }
        Start-Sleep -Milliseconds 250
    }

    $remainingOut = & $processNewLines -Path $stdoutLog -Cursor ([ref]$outLineCount)
    $remainingErr = & $processNewLines -Path $stderrLog -Cursor ([ref]$errLineCount)
    $remaining = @($remainingOut + $remainingErr)
    foreach ($line in $remaining) {
        if ($line -ne $null -and $line -ne '') {
            Write-Host $line
            if (-not $qrOpened -and ($line -match $regex)) {
                $qrOpened = $true
                Open-AccessQrPage -LanUrl $lanUrl -TunnelUrl $Matches[0]
            }
        }
    }

    if (-not $qrOpened) {
        Write-Host "[WARN] URL tunnel tidak terdeteksi. Membuka QR lokal saja." -ForegroundColor Yellow
        Open-AccessQrPage -LanUrl $lanUrl
    }

    Remove-Item -LiteralPath $stdoutLog -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stderrLog -ErrorAction SilentlyContinue

    if ($process.ExitCode -ne 0) {
        if ((Test-Path -LiteralPath $stderrLog) -and (Get-Content -LiteralPath $stderrLog -ErrorAction SilentlyContinue | Select-String -SimpleMatch 'failed to unmarshal quick Tunnel')) {
            Write-Host "[ERROR] Cloudflare quick tunnel gagal diproses oleh cloudflared." -ForegroundColor Red
            Write-Host "Kemungkinan penyebab: cloudflared terlalu lama, koneksi/proxy mengubah respon Cloudflare, atau akses keluar diblokir." -ForegroundColor Yellow
            Write-Host "Coba update cloudflared ke versi terbaru, lalu jalankan lagi." -ForegroundColor Yellow
        }
        throw "cloudflared berhenti dengan exit code $($process.ExitCode)."
    }
}
catch {
    Write-Host "Terjadi error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Cara jalankan yang aman dari PowerShell:" -ForegroundColor Yellow
    Write-Host "powershell -ExecutionPolicy Bypass -File `"$PSScriptRoot\Online Local Bun.ps1`"" -ForegroundColor White
    Write-Host "Atau tanpa tunnel:" -ForegroundColor Yellow
    Write-Host "powershell -ExecutionPolicy Bypass -File `"$PSScriptRoot\Online Local Bun.ps1`" -NoTunnel" -ForegroundColor White
}
finally {
    Pause-IfNeeded
}
