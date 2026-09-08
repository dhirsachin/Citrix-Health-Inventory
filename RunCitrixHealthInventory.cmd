@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem Launches the saved script in a clean Windows PowerShell 5.1 process.
rem -ExecutionPolicy Bypass applies only to this child process; it does not
rem change LocalMachine or CurrentUser execution-policy settings.
set "quiet_mode="
if /I "%~1"=="/quiet" set "quiet_mode=1"
set "script_path=%~dp0CitrixHealthInventory.ps1"
set "powershell_path=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "launcher_log_dir=%LOCALAPPDATA%\CitrixDataPull\Logs"
set "launcher_log=%launcher_log_dir%\CitrixHealthInventory_Launcher.log"

if not exist "%launcher_log_dir%" mkdir "%launcher_log_dir%" >nul 2>&1
>>"%launcher_log%" echo [%date% %time%] Launcher started. Script="%script_path%"

if not exist "%script_path%" (
    >>"%launcher_log%" echo [%date% %time%] ERROR Required script was not found: "%script_path%"
    echo Required script was not found: "%script_path%"
    if not defined quiet_mode pause
    exit /b 2
)

if not exist "%powershell_path%" (
    >>"%launcher_log%" echo [%date% %time%] ERROR Windows PowerShell 5.1 executable was not found: "%powershell_path%"
    echo Windows PowerShell 5.1 executable was not found: "%powershell_path%"
    if not defined quiet_mode pause
    exit /b 3
)

"%powershell_path%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%script_path%" >>"%launcher_log%" 2>&1
set "launch_exit=%ERRORLEVEL%"
>>"%launcher_log%" echo [%date% %time%] Launcher completed. ExitCode=%launch_exit%

if not "%launch_exit%"=="0" (
    echo Citrix Health and Inventory failed with exit code %launch_exit%.
    echo Launcher details: "%launcher_log%"
    if not defined quiet_mode pause
)
exit /b %launch_exit%
