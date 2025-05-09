# PowerShell System Monitor Logger with LibreHardwareMonitor integration

$logFile = "system_log.csv"
$lhmPath = "C:\Tools\LibreHardwareMonitor\LibreHardwareMonitor.exe"
$lhmLogPath = "C:\Tools\lhm_logs\latest.csv"
$targetProcess = "MonsterHunterWilds"
$checkInterval = 5

# Launch LibreHardwareMonitor if not already running
if (-not (Get-Process | Where-Object { $_.Path -eq $lhmPath })) {
    Start-Process -FilePath $lhmPath -WindowStyle Minimized
    Start-Sleep -Seconds 3
}

# Initialize log
if (!(Test-Path $logFile)) {
    "Timestamp,CPU_Temp,CPU_Usage,GPU_Temp,GPU_Usage,RAM_Usage,Disk_IO,Net_Down,Net_Up" | Out-File $logFile
}

function Get-TempFromLHM {
    if (!(Test-Path $lhmLogPath)) {
        return @{ CPU_Temp = 0; GPU_Temp = 0 }
    }

    $lines = Get-Content $lhmLogPath
    $lastLine = $lines[-1]
    $columns = $lastLine -split ','

    $cpuTemp = 0
    $gpuTemp = 0

    foreach ($line in $lines) {
        if ($line -match "CPU Package" -and $line -match "Temperature") {
            $cpuTemp = [double]($line -split ',')[-1]
        }
        if ($line -match "GPU Core" -and $line -match "Temperature") {
            $gpuTemp = [double]($line -split ',')[-1]
        }
    }

    return @{ CPU_Temp = $cpuTemp; GPU_Temp = $gpuTemp }
}

function Get-SystemStats {
    $temps = Get-TempFromLHM
    $stats = @{
        Timestamp = (Get-Date).ToString("s")
        CPU_Temp = $temps.CPU_Temp
        CPU_Usage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples[0].CookedValue
        GPU_Temp = $temps.GPU_Temp
        GPU_Usage = (Get-Counter '\GPU Engine(*)\Utilization Percentage' | Where-Object { $_.InstanceName -like "*engtype_3D*" }).CounterSamples | Measure-Object -Property CookedValue -Average | Select-Object -ExpandProperty Average
        RAM_Usage = (Get-Counter '\Memory\% Committed Bytes In Use').CounterSamples[0].CookedValue
        Disk_IO = (Get-Counter '\PhysicalDisk(_Total)\Disk Bytes/sec').CounterSamples[0].CookedValue / 1MB
        Net_Down = (Get-Counter '\Network Interface(*)\Bytes Received/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum
        Net_Up = (Get-Counter '\Network Interface(*)\Bytes Sent/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum
    }
    return $stats
}

function Write-StatsToCSV($stats) {
    "$($stats.Timestamp),$([math]::Round($stats.CPU_Temp,2)),$([math]::Round($stats.CPU_Usage,2)),$([math]::Round($stats.GPU_Temp,2)),$([math]::Round($stats.GPU_Usage,2)),$([math]::Round($stats.RAM_Usage,2)),$([math]::Round($stats.Disk_IO,2)),$([math]::Round($stats.Net_Down / 1MB,2)),$([math]::Round($stats.Net_Up / 1MB,2))" | Out-File -FilePath $logFile -Append
}

$wasRunning = $false

while ($true) {
    $isRunning = Get-Process -Name $targetProcess -ErrorAction SilentlyContinue

    $stats = Get-SystemStats
    Write-StatsToCSV $stats

    if ($isRunning) {
        if (-not $wasRunning) {
            Write-Host "Session started: $targetProcess"
        }
        $wasRunning = $true
    } else {
        if ($wasRunning) {
            Write-Host "Session ended: $targetProcess"
            Start-Sleep -Seconds 2

            # Run Excel macro after session ends
            $excelPath = "C:\Path\To\system_monitor_template.xlsm"
            $macroName = "AutoGraphSystemData"

            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($excelPath)
            $excel.Run($macroName)
            $workbook.Save()
            $workbook.Close($false)
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()

            $wasRunning = $false
        }
    }

    Start-Sleep -Seconds $checkInterval
}