# PowerShell: log_system_monitor.ps1

# Display message box when the script starts
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("System Monitor Logger is now running...", "Logger Started", "OK", "Information")

# Define new log directory path
$logPath = "C:\Tools\logs"

# Ensure log directory exists BEFORE using it
if (!(Test-Path $logPath)) {
    try {
        Write-Output "Creating log directory: $logPath"
        New-Item -ItemType Directory -Path $logPath -Force | Out-Null
        Start-Sleep -Milliseconds 500  # brief pause to ensure file system has updated
    } catch {
        Write-Error "Failed to create log directory at $logPath. Error: $_"
        exit
    }
}

# Confirm directory is accessible
if (-not (Test-Path -Path $logPath -IsValid)) {
    Write-Error "Log path is invalid or not accessible: $logPath"
    exit
}

# Create timestamped CSV file path (after confirming logPath exists)
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvFile = Join-Path $logPath "session_$timestamp.csv"

# Target game process
$targetProcess = "MonsterHunterWilds.exe"

# Write CSV header
try {
    Write-Output "Creating CSV file: $csvFile"
    "Timestamp,CPU_Usage,Memory_Used(MB),Memory_Total(MB),Disk_Read_Bps,Disk_Write_Bps,Net_Received_Bps,Net_Sent_Bps,ActiveProcesses" | Out-File -FilePath $csvFile -Encoding utf8
} catch {
    Write-Error "Could not create CSV file: $csvFile. Error: $_"
    exit
}

# Game session tracking
$gameStarted = $false
$sessionRunning = $true
Write-Output "Monitoring session started..."

# Start monitoring loop
while ($sessionRunning) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Get system metrics
    $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
    $os = Get-CimInstance Win32_OperatingSystem
    $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1024, 2)
    $memFree = [math]::Round($os.FreePhysicalMemory / 1024, 2)
    $memUsed = $memTotal - $memFree
    $diskRead = (Get-Counter '\PhysicalDisk(_Total)\Disk Read Bytes/sec').CounterSamples.CookedValue
    $diskWrite = (Get-Counter '\PhysicalDisk(_Total)\Disk Write Bytes/sec').CounterSamples.CookedValue
    $netRecv = (Get-Counter '\Network Interface(*)\Bytes Received/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum
    $netSent = (Get-Counter '\Network Interface(*)\Bytes Sent/sec').CounterSamples | Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum
    $activeProcs = (Get-Process | Select-Object -ExpandProperty ProcessName) -join ";"

    # Check if the game has started
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

    # Log system data to CSV
    try {
        "$timestamp,$cpu,$memUsed,$memTotal,$diskRead,$diskWrite,$netRecv,$netSent,""$activeProcs""" | Out-File -Append -FilePath $csvFile -Encoding utf8
    } catch {
        Write-Error "Failed to append to CSV: $csvFile. Error: $_"
        exit
    }

    Start-Sleep -Seconds 5
}

# Open Excel and run macro with the CSV file path as an argument
$excelPath = "C:\Users\Miguel\OneDrive\Documents\GitHub\monsterhunterloggerdogger\excel\system_monitor_template.xlsm"

if (Test-Path $excelPath) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelPath)

    # Run macro with CSV path (you must update your VBA macro to accept and import this CSV)
    try {
        $excel.Run("AutoGraphSystemData", $csvFile)
        $workbook.Save()
        Write-Output "Excel macro executed to generate graphs."
    } catch {
        Write-Warning "Failed to run Excel macro: $_"
    }

    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} else {
    Write-Warning "Excel template not found at $excelPath"
}
