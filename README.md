# Real-Time Excel Vehicle Expiration Notification System

This system monitors an Excel file (`VehicleMonitoring.xlsx`) for vehicle plate expirations and gives real-time Windows popup notifications based on status changes.

## Prerequisites
1. **Python 3.8+** must be installed on your Windows system.
   - When installing Python, **ensure you check the box that says "Add Python to PATH"** at the bottom of the installer window.
   - You can download Python from [python.org/downloads](https://www.python.org/downloads/).

## Setup Instructions

1. **Install Dependencies**
   Open your command prompt or PowerShell in this folder (`c:\Users\gitga\Documents\vehicle expiration alert`) and run:
   ```cmd
   pip install -r requirements.txt
   ```

2. **Generate Mock Data (Optional)**
   If you don't already have an Excel file named `VehicleMonitoring.xlsx`, you can generate an example file by running:
   ```cmd
   python generate_mock_data.py
   ```

3. **Run the Monitor**
   Start the monitoring script by running:
   ```cmd
   python vehicle_monitor.py
   ```
   The script will start, perform an initial scan of the file, and continue running in the background. It will output logs to the console and show a Windows Toast Notification when expiration statuses change.

## Features
- **Permanent System Tray App**: Minimizes to the taskbar area near the clock. It runs silently in the background!
- **Manual Scanning Dropdown**: Right-click the system tray icon to access a "Scan All" option, or choose to scan a specific Month/Sheet manually from a dropdown!
- **Real-Time Monitoring**: Automatically detects changes when the Excel file is modified and saved.
- **Smart Memory**: Remembers last scanned statuses to prevent duplicate toast notifications.

## Creating the Executable (.exe)
If you want to create a portable `.exe` version of this application that you can set to auto-start with Windows:

1. Double click the `build.bat` file in this directory.
2. Wait for PyInstaller to finish bundling the Python environment.
3. Open the newly created `dist/` folder.
4. Your permanent app will be `vehicle_monitor.exe`.
5. To make it start automatically, press `WIN + R`, type `shell:startup`, and drag a **Shortcut** of `vehicle_monitor.exe` into that folder!
