param(
    [string]$Port
)

function Get-AutoSerialPort {
    param(
        [string]$PreferredPort
    )

    if ($PreferredPort) {
        return $PreferredPort
    }

    $devices = @(Get-CimInstance Win32_SerialPort -ErrorAction SilentlyContinue)
    if (-not $devices) {
        throw "No serial ports found. Connect the ESP32-S3 and try again."
    }

    $candidates = foreach ($device in $devices) {
        $text = @($device.Name, $device.Description, $device.PNPDeviceID) -join ' '
        $score = 0

        if ($text -match 'VID_303A|ESP32|Espressif|USB JTAG') { $score += 100 }
        if ($text -match 'VID_10C4|CP210|Silicon Labs') { $score += 60 }
        if ($text -match 'VID_1A86|CH340|wch') { $score += 40 }
        if ($text -match 'USB Serial|UART') { $score += 15 }

        [PSCustomObject]@{
            Port = $device.DeviceID
            Name = $device.Name
            Score = $score
        }
    }

    $ranked = @($candidates | Sort-Object Score, Port -Descending)
    if ($ranked.Count -eq 1 -or $ranked[0].Score -gt $ranked[1].Score) {
        Write-Host "=== Auto-detected ESP32 serial port: $($ranked[0].Port) [$($ranked[0].Name)] ==="
        return $ranked[0].Port
    }

    $details = ($ranked | ForEach-Object { "  $($_.Port) - $($_.Name)" }) -join [Environment]::NewLine
    throw "Unable to choose a single ESP32 serial port automatically. Use -Port <COMx>. Candidates:`n$details"
}

# Remove MSYS environment variables that trigger ESP-IDF's MinGW rejection
Remove-Item env:MSYSTEM -ErrorAction SilentlyContinue
Remove-Item env:MSYSTEM_CARCH -ErrorAction SilentlyContinue
Remove-Item env:MSYSTEM_CHOST -ErrorAction SilentlyContinue
Remove-Item env:MSYSTEM_PREFIX -ErrorAction SilentlyContinue
Remove-Item env:MINGW_CHOST -ErrorAction SilentlyContinue
Remove-Item env:MINGW_PACKAGE_PREFIX -ErrorAction SilentlyContinue
Remove-Item env:MINGW_PREFIX -ErrorAction SilentlyContinue

if (-not $env:IDF_PATH) {
    $env:IDF_PATH = "C:\Users\ruv\esp\v5.4\esp-idf"
}
if (-not $env:IDF_TOOLS_PATH) {
    $env:IDF_TOOLS_PATH = "C:\Espressif\tools"
}
if (-not $env:IDF_PYTHON_ENV_PATH) {
    $env:IDF_PYTHON_ENV_PATH = "C:\Espressif\tools\python\v5.4\venv"
}

$idfToolPaths = @(
    "C:\Espressif\tools\xtensa-esp-elf\esp-14.2.0_20241119\xtensa-esp-elf\bin",
    "C:\Espressif\tools\cmake\3.30.2\cmake-3.30.2-windows-x86_64\bin",
    "C:\Espressif\tools\ninja\1.12.1",
    "C:\Espressif\tools\ccache\4.10.2\ccache-4.10.2-windows-x86_64",
    "C:\Espressif\tools\idf-exe\1.0.3",
    "C:\Espressif\tools\python\v5.4\venv\Scripts"
) | Where-Object { Test-Path $_ }

if ($idfToolPaths) {
    $env:PATH = (($idfToolPaths -join ';') + ';' + $env:PATH)
}

Set-Location $PSScriptRoot

$python = "$env:IDF_PYTHON_ENV_PATH\Scripts\python.exe"
$idf = "$env:IDF_PATH\tools\idf.py"

Write-Host "=== Cleaning stale build cache ==="
& $python $idf fullclean

Write-Host "=== Building firmware (SSID=ruv.net, target=192.168.1.20:5005) ==="
& $python $idf build

if ($LASTEXITCODE -eq 0) {
    $resolvedPort = Get-AutoSerialPort -PreferredPort $Port
    Write-Host "=== Build succeeded! Flashing to $resolvedPort ==="
    & $python $idf -p $resolvedPort flash
} else {
    Write-Host "=== Build failed with exit code $LASTEXITCODE ==="
}
