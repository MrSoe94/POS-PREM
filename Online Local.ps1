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

function Open-QrForTunnelUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TunnelUrl
    )
    $encoded = [uri]::EscapeDataString($TunnelUrl)
    $qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=$encoded"
    Write-Host "Link tunnel: $TunnelUrl" -ForegroundColor Green
    Write-Host "Membuka QR Code di browser..." -ForegroundColor Cyan
    Start-Process $qrUrl | Out-Null
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

try {
    Write-Host "Memulai Aplikasi..." -ForegroundColor Cyan
    Write-Host "Lokasi project: $PSScriptRoot" -ForegroundColor DarkGray

    if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
        throw "npm tidak ditemukan. Pastikan Node.js sudah terpasang dan ada di PATH."
    }

    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'server.js'))) {
        throw "server.js tidak ditemukan di folder project. Jalankan script dari folder aplikasi yang benar."
    }

    # Jalankan npm start di jendela PowerShell terpisah
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", "npm start") `
        -WorkingDirectory $PSScriptRoot |
        Out-Null

    Write-Host "Menunggu server aktif..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    if ($NoTunnel) {
        Write-Host "Mode lokal aktif. Buka: http://localhost:$Port" -ForegroundColor Green
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
                    Open-QrForTunnelUrl -TunnelUrl $Matches[0]
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
                Open-QrForTunnelUrl -TunnelUrl $Matches[0]
            }
        }
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
    Write-Host "powershell -ExecutionPolicy Bypass -File `"$PSScriptRoot\Online Local.ps1`"" -ForegroundColor White
    Write-Host "Atau tanpa tunnel:" -ForegroundColor Yellow
    Write-Host "powershell -ExecutionPolicy Bypass -File `"$PSScriptRoot\Online Local.ps1`" -NoTunnel" -ForegroundColor White
}
finally {
    Pause-IfNeeded
}
