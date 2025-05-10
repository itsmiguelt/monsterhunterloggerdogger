# Path to LibreHardwareMonitor DLL
$librePath = "C:\Tools\LibreHardwareMonitor-net472\LibreHardwareMonitorLib.dll"

# Validate DLL path
if (-not (Test-Path $librePath)) {
    Write-Error "LibreHardwareMonitorLib.dll not found at $librePath"
    exit 1
}

# Log path configuration
$logPath = "C:\Tools\logs"
try {
    if (!(Test-Path $logPath)) {
        New-Item -ItemType Directory -Path $logPath -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Error "Failed to create log directory at $logPath : $_"
    exit 1
}

# Load DLL with better error handling
try {
    Add-Type -Path $librePath -ErrorAction Stop
    Write-Output "Successfully loaded LibreHardwareMonitorLib.dll"
}
catch {
    Write-Error "Failed to load DLL: $_"
    if ($_.Exception.Message -like "*bad format*") {
        Write-Warning "This might be a 32-bit/64-bit mismatch. Ensure PowerShell and DLL architectures match."
    }
    exit 1
}

# Instantiate computer object with all sensors enabled
$computer = New-Object LibreHardwareMonitor.Hardware.Computer -ErrorAction Stop
$computer.IsCpuEnabled = $true
$computer.IsGpuEnabled = $true
$computer.IsMemoryEnabled = $true
$computer.IsMotherboardEnabled = $true
$computer.IsControllerEnabled = $true
$computer.IsNetworkEnabled = $true
$computer.IsStorageEnabled = $true

try {
    $computer.Open()
}
catch {
    Write-Error "Failed to initialize hardware monitoring: $_"
    exit 1
}

# Sensor dump for debugging
$sensorDumpPath = Join-Path $logPath "sensor_dump_$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
try {
    "Sensor Dump - $(Get-Date)" | Out-File -FilePath $sensorDumpPath -Encoding utf8 -ErrorAction Stop
    
    foreach ($hw in $computer.Hardware) {
        $hw.Update()
        foreach ($sub in $hw.SubHardware) { 
            try { $sub.Update() } catch { Write-Warning "Failed to update subhardware: $_" }
        }
        foreach ($sensor in $hw.Sensors) {
            "$($hw.HardwareType) - $($sensor.Name) - $($sensor.SensorType) - $($sensor.Value)" | 
                Out-File -Append -FilePath $sensorDumpPath -Encoding utf8
        }
    }
}
catch {
    Write-Warning "Failed to create sensor dump: $_"
}

# Verify sensors were found
if ($computer.Hardware.Count -eq 0) {
    Write-Error "No hardware devices detected. Possible causes:`n" +
                "1. Running without administrator privileges`n" +
                "2. Hardware not supported by LibreHardwareMonitor`n" +
                "3. Driver issues with your hardware"
    exit 1
}

# Create timestamped CSV file
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvFile = Join-Path $logPath "session_$timestamp.csv"
$targetProcess = "MonsterHunterWilds.exe"

# Write CSV header
try {
    "Timestamp,CPU_Usage,CPU_Temp,GPU_Usage,GPU_Temp,Memory_Used(MB),Memory_Total(MB),Disk_Read_Bps,Disk_Write_Bps,Net_Received_Bps,Net_Sent_Bps,ActiveProcesses" | 
        Out-File -FilePath $csvFile -Encoding utf8 -ErrorAction Stop
}
catch {
    Write-Error "Could not create CSV file: $csvFile`n$_"
    exit 1
}

function Get-SensorValue {
    param (
        [string]$hardwareType,
        [string]$sensorType,
        [string]$namePattern
    )
    
    try {
        foreach ($hw in $computer.Hardware) {
            if ($hw.HardwareType -eq $hardwareType) {
                try { $hw.Update() } catch { Write-Debug "Update failed for $($hw.Name)" }
                
                foreach ($sub in $hw.SubHardware) { 
                    try { $sub.Update() } catch { Write-Debug "Subhardware update failed" }
                }
                
                $matchingSensors = $hw.Sensors | Where-Object {
                    $_.SensorType -eq $sensorType -and $_.Name -like "*$namePattern*"
                }
                
                if ($matchingSensors) {
                    return [math]::Round(($matchingSensors | Select-Object -First 1).Value, 2)
                }
            }
        }
        
        # Try alternative naming patterns if first attempt fails
        if ($hardwareType -eq "Cpu" -and $sensorType -eq "Temperature") {
            $altPatterns = @("Package", "Core", "CPU", "Tdie", "Tctl")
            foreach ($pattern in $altPatterns) {
                $result = Get-SensorValue -hardwareType $hardwareType -sensorType $sensorType -namePattern $pattern
                if ($result -ne $null) { return $result }
            }
        }
        
        return $null
    }
    catch {
        Write-Debug "Sensor read error: $_"
        return $null
    }
}

$gameStarted = $false
$sessionRunning = $true
Write-Output "Monitoring session started. Press Ctrl+C to stop."

try {
    while ($sessionRunning) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Get CPU metrics with fallback
        $cpu = "NA"
        try {
            $cpuCounter = Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction SilentlyContinue
            if ($cpuCounter) {
                $cpu = [math]::Round($cpuCounter.CounterSamples.CookedValue, 2)
            }
        } catch {
            Write-Warning "Failed to get CPU usage: $_"
        }
        
        $cpuTemp = Get-SensorValue -hardwareType "Cpu" -sensorType "Temperature" -namePattern "Package"
        if ($null -eq $cpuTemp) { $cpuTemp = "NA" }
        
        # Get GPU metrics (try both NVIDIA and AMD)
        $gpuUsage = Get-SensorValue -hardwareType "GpuNvidia" -sensorType "Load" -namePattern "Core"
        $gpuTemp = Get-SensorValue -hardwareType "GpuNvidia" -sensorType "Temperature" -namePattern "GPU Core"
        
        if ($null -eq $gpuUsage) {
            $gpuUsage = Get-SensorValue -hardwareType "GpuAmd" -sensorType "Load" -namePattern "Core"
            $gpuTemp = Get-SensorValue -hardwareType "GpuAmd" -sensorType "Temperature" -namePattern "Core"
        }
        
        if ($null -eq $gpuUsage) { $gpuUsage = "NA" }
        if ($null -eq $gpuTemp) { $gpuTemp = "NA" }
        
        # Get memory metrics
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1024, 2)
            $memFree = [math]::Round($os.FreePhysicalMemory / 1024, 2)
            $memUsed = $memTotal - $memFree
        } catch {
            $memUsed = "NA"
            $memTotal = "NA"
            Write-Warning "Failed to get memory info: $_"
        }
        
        # Get disk metrics with fallback
        $diskRead = "NA"
        $diskWrite = "NA"
        try {
            $diskRead = (Get-Counter '\PhysicalDisk(_Total)\Disk Read Bytes/sec' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            $diskWrite = (Get-Counter '\PhysicalDisk(_Total)\Disk Write Bytes/sec' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            if ($diskRead) { $diskRead = [math]::Round($diskRead, 2) }
            if ($diskWrite) { $diskWrite = [math]::Round($diskWrite, 2) }
        } catch {
            Write-Warning "Failed to get disk metrics: $_"
        }
        
        # Get network metrics with proper null handling
        $netRecv = "NA"
        $netSent = "NA"
        try {
            $netRecvCounter = Get-Counter '\Network Interface(*)\Bytes Received/sec' -ErrorAction SilentlyContinue
            $netSentCounter = Get-Counter '\Network Interface(*)\Bytes Sent/sec' -ErrorAction SilentlyContinue
            
            if ($netRecvCounter) {
                $sum = ($netRecvCounter.CounterSamples | Measure-Object -Property CookedValue -Sum).Sum
                if ($sum -ne $null) { $netRecv = [math]::Round($sum, 2) }
            }
            
            if ($netSentCounter) {
                $sum = ($netSentCounter.CounterSamples | Measure-Object -Property CookedValue -Sum).Sum
                if ($sum -ne $null) { $netSent = [math]::Round($sum, 2) }
            }
        } catch {
            Write-Warning "Failed to get network metrics: $_"
        }
        
        # Get process list
        try {
            $activeProcs = ((Get-Process -ErrorAction SilentlyContinue).ProcessName | Sort-Object -Unique) -join ";"
        } catch {
            $activeProcs = "NA"
            Write-Warning "Failed to get process list: $_"
        }
        
        # Check for target process
        if (Get-Process -Name ($targetProcess -replace ".exe$", "") -ErrorAction SilentlyContinue) {
            if (-not $gameStarted) {
                Write-Output "Game detected at $timestamp"
                $gameStarted = $true
            }
        } elseif ($gameStarted) {
            Write-Output "Game closed. Ending session."
            $sessionRunning = $false
            break
        }
        
        # Write to CSV
        try {
            "$timestamp,$cpu,$cpuTemp,$gpuUsage,$gpuTemp,$memUsed,$memTotal,$diskRead,$diskWrite,$netRecv,$netSent,""$activeProcs""" | 
                Out-File -Append -FilePath $csvFile -Encoding utf8 -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to write to CSV: $_"
            $sessionRunning = $false
            break
        }
        
        Start-Sleep -Seconds 5
    }
}
finally {
    # Cleanup
    try { $computer.Close() } catch { Write-Warning "Failed to clean up hardware monitor: $_" }
    
    # Auto-run Excel macro if configured
    $excelPath = "C:\Users\Miguel\OneDrive\Documents\GitHub\monsterhunterloggerdogger\excel\system_monitor_template.xlsm"
    if (Test-Path $excelPath) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open($excelPath)
            $excel.Run("AutoGraphSystemData", $csvFile)
            $workbook.Save()
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Write-Output "Excel report generated successfully."
        }
        catch {
            Write-Warning "Failed to generate Excel report: $_"
        }
    }
    else {
        Write-Warning "Excel template not found at $excelPath"
    }
    
    Write-Output "Monitoring session completed. Data saved to $csvFile"
}