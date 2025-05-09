# Monster Hunter Logger Dogger

Logs and visualizes PC performance during gaming sessions using PowerShell and LibreHardwareMonitor.

## Features

- Monitors CPU/GPU temperatures and usage.
- Logs RAM, disk I/O, and network activity.
- Detects when Monster Hunter Wilds starts and ends.
- Automatically generates Excel graphs post-session.

## Setup

1. **Install LibreHardwareMonitor** and configure it to log data to `C:\Tools\lhm_logs\latest.csv`.
2. **Place `system_monitor_template.xlsm`** in `C:\Tools\`.
3. **Run the PowerShell script**:
   ```powershell
   .\scripts\log_system_monitor.ps1
