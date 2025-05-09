# PowerShell: log_system_monitor.ps1
$logPath = "$PSScriptRoot\..\logs"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvFile = Join-Path $logPath "session_$timestamp.csv"
$targetProcess = "MonsterHunterWilds.exe"

if (!(Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath | Out-Null
}

# Header
"Timestamp,CPU_Usage,Memory_Used(MB),Memory_Total(MB),Disk_Read_Bps,Disk_Write_Bps,Net_Received_Bps,Net_Sent_Bps,ActiveProcesses" | Out-File -FilePath $csvFile -Encoding utf8

# Check for game session start
$gameStarted = $false
$sessionRunning = $true

Write-Output "Monitoring session started..."

while ($sessionRunning) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Get CPU Usage
    $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue

    # Get RAM Usage
    $os = Get-CimInstance Win32_OperatingSystem
    $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1024, 2)
    $memFree = [math]::Round($os.FreePhysicalMemory / 1024, 2)
    $memUsed = $memTotal - $memFree

    # Get Disk Read/Write
    $diskRead = (Get-Counter '\PhysicalDisk(_Total)\Disk Read Bytes/sec').CounterSamples.CookedValue
    $diskWrite = (Get-Counter '\PhysicalDisk(_Total)\Disk Write Bytes/sec').CounterSamples.CookedValue

    # Get Network Usage
    $netRecv = (Get-Counter '\Network Interface(*)\Bytes Received/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum
    $netSent = (Get-Counter '\Network Interface(*)\Bytes Sent/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum

    # Get process list
    $activeProcs = (Get-Process | Select-Object -ExpandProperty ProcessName) -join ";"

    # Check for game start
    if (Get-Process -Name ($targetProcess -replace ".exe$", "") -ErrorAction SilentlyContinue) {
        if (-not $gameStarted) {
            Write-Output "Game started at $timestamp"
            $gameStarted = $true
        }
    } elseif ($gameStarted) {
        Write-Output "Game closed. Ending session."
        $sessionRunning = $false
        break
    }

    # Log to CSV
    "$timestamp,$cpu,$memUsed,$memTotal,$diskRead,$diskWrite,$netRecv,$netSent,""$activeProcs""" | Out-File -Append -FilePath $csvFile -Encoding utf8

    Start-Sleep -Seconds 5
}

# Open Excel and run macro
$excelPath = "$PSScriptRoot\system_monitor_template.xlsm"
if (Test-Path $excelPath) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelPath)
    $excel.Run("GenerateGraphs", $csvFile)
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Output "Excel macro executed to generate graphs."
} else {
    Write-Warning "Excel template not found at $excelPath"
}