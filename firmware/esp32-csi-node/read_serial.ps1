param(
    [string]$Port,
    [int]$Baud = 115200
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

$resolvedPort = Get-AutoSerialPort -PreferredPort $Port
$p = New-Object System.IO.Ports.SerialPort($resolvedPort, $Baud)
$p.ReadTimeout = 5000
$p.Open()
Start-Sleep -Milliseconds 200

for ($i = 0; $i -lt 60; $i++) {
    try {
        $line = $p.ReadLine()
        Write-Host $line
    } catch {
        break
    }
}
$p.Close()
