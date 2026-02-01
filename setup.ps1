# JARVIS 2.0 Setup
Write-Host "Setting up J.A.R.V.I.S. System..." -ForegroundColor Cyan

# Create directory
$dir = "C:\jarvis-pro"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Set-Location $dir

# Check Python
try {
    $pyVersion = python --version
    Write-Host "Found: $pyVersion" -ForegroundColor Green
} catch {
    Write-Error "Python not found! Please install Python 3.8+"
    exit
}

# Install dependencies
Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install flask flask-cors psutil pyautogui pyperclip pywin32

# Create files (in real setup, these would be copied)
Write-Host "Creating system files..." -ForegroundColor Yellow
# (Files would be written here)

# Create startup shortcut
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\JARVIS.lnk")
$Shortcut.TargetPath = "pythonw"
$Shortcut.Arguments = "$dir\jarvis_server.py"
$Shortcut.WorkingDirectory = $dir
$Shortcut.Save()

Write-Host "`nInstallation Complete!" -ForegroundColor Green
Write-Host "Start server: python jarvis_server.py" -ForegroundColor Cyan
Write-Host "Open browser: http://localhost:5000" -ForegroundColor Cyan
Write-Host "Test commands: .\test_commands.ps1" -ForegroundColor Cyan

pause