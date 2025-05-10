<#
.SYNOPSIS
Monitors system performance while MonsterHunterWilds is running and generates Excel report

.DESCRIPTION
Logs system metrics to CSV and auto-generates Excel report with graphs when game closes
#>

#Requires -RunAsAdministrator

# Configuration
$librePath = "C:\Tools\LibreHardwareMonitor-net472\LibreHardwareMonitorLib.dll"
$logPath = "C:\Tools\logs"
$targetProcess = "MonsterHunterWilds"
$pollingInterval = 5 # seconds
$excelPath = "C:\Users\Miguel\OneDrive\Documents\GitHub\monsterhunterloggerdogger\excel\system_monitor_template.xlsm"

# Create log directory if needed
if (!(Test-Path $logPath)) { 
    try {
        New-Item -ItemType Directory -Path $logPath -Force | Out-Null
    }
    catch {
        Write-Error "Failed to create log directory: $_"
        exit 1
    }
}

# Wait for game to fully initialize
Write-Output "Waiting for game to initialize..."
Start-Sleep -Seconds 20

# Verify game is actually running
if (-not (Get-Process $targetProcess -ErrorAction SilentlyContinue)) {
    Write-Output "Game process not found - exiting"
    exit
}

# Load LibreHardwareMonitor
try {
    Add-Type -Path $librePath
    $computer = New-Object LibreHardwareMonitor.Hardware.Computer
    $computer.IsCpuEnabled = $true
    $computer.IsGpuEnabled = $true
    $computer.IsMemoryEnabled = $true
    $computer.IsStorageEnabled = $true
    $computer.Open()
}
catch {
    Write-Error "Failed to initialize hardware monitoring: $_"
    exit 1
}

# Create CSV log file
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$csvFile = Join-Path $logPath "mhw_metrics_$timestamp.csv"

"Timestamp,CPU_Usage(%),CPU_Temp(°C),GPU_Usage(%),GPU_Temp(°C),RAM_Used(GB),RAM_Total(GB),Disk_Read(MB/s),Disk_Write(MB/s)" | Out-File $csvFile

function Get-SensorValue {
    param($hardwareType, $sensorType, $namePattern)
    
    foreach ($hw in $computer.Hardware) {
        if ($hw.HardwareType -eq $hardwareType) {
            $hw.Update()
            $sensor = $hw.Sensors | Where-Object { 
                $_.SensorType -eq $sensorType -and $_.Name -like "*$namePattern*" 
            } | Select-Object -First 1
            if ($sensor) { return [math]::Round($sensor.Value, 2) }
        }
    }
    return $null
}

Write-Output "Starting monitoring for $targetProcess..."
try {
    while ($true) {
        # Exit if game closed
        if (-not (Get-Process $targetProcess -ErrorAction SilentlyContinue)) {
            Write-Output "Game closed - stopping monitoring"
            break
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Get CPU metrics
        $cpuUsage = try { [math]::Round((Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue, 2) } catch { "NA" }
        $cpuTemp = Get-SensorValue -hardwareType "Cpu" -sensorType "Temperature" -namePattern "Package"
        if ($null -eq $cpuTemp) { $cpuTemp = "NA" }
        
        # Get GPU metrics
        $gpuUsage = Get-SensorValue -hardwareType "GpuNvidia" -sensorType "Load" -namePattern "Core"
        $gpuTemp = Get-SensorValue -hardwareType "GpuNvidia" -sensorType "Temperature" -namePattern "Core"
        
        if ($null -eq $gpuUsage) {
            $gpuUsage = Get-SensorValue -hardwareType "GpuAmd" -sensorType "Load" -namePattern "Core"
            $gpuTemp = Get-SensorValue -hardwareType "GpuAmd" -sensorType "Temperature" -namePattern "Core"
        }
        if ($null -eq $gpuUsage) { $gpuUsage = "NA" }
        if ($null -eq $gpuTemp) { $gpuTemp = "NA" }

        # Get memory metrics
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $memUsed = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 2)
            $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        } catch {
            $memUsed = "NA"
            $memTotal = "NA"
        }

        # Get disk metrics
        $diskRead = try { [math]::Round((Get-Counter '\PhysicalDisk(_Total)\Disk Read Bytes/sec' -ErrorAction Stop).CounterSamples.CookedValue / 1MB, 2) } catch { "NA" }
        $diskWrite = try { [math]::Round((Get-Counter '\PhysicalDisk(_Total)\Disk Write Bytes/sec' -ErrorAction Stop).CounterSamples.CookedValue / 1MB, 2) } catch { "NA" }

        # Write to CSV
        "$timestamp,$cpuUsage,$cpuTemp,$gpuUsage,$gpuTemp,$memUsed,$memTotal,$diskRead,$diskWrite" | Out-File -Append $csvFile
        
        Start-Sleep -Seconds $pollingInterval
    }
}
finally {
    try {
        $computer.Close()
        Write-Output "Monitoring stopped. Data saved to $csvFile"
        
        # Generate Excel report
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
    }
    catch {
        Write-Warning "Error during cleanup: $_"
    }
}