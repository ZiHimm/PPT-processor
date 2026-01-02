# Enhanced Marketing Dashboard Automator Launcher (PowerShell)
Write-Host "Marketing Dashboard Automator v2.0" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

# Find Python 3.11
$pythonPaths = @(
    "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python311\python.exe",
    "C:\Python311\python.exe",
    "C:\Program Files\Python311\python.exe",
    "C:\Program Files (x86)\Python311\python.exe"
)

$pythonExe = $null
foreach ($path in $pythonPaths) {
    if (Test-Path $path) {
        $pythonExe = $path
        Write-Host "Found Python 3.11 at: $pythonExe" -ForegroundColor Green
        break
    }
}

if (-not $pythonExe) {
    Write-Host "Python 3.11 not found!" -ForegroundColor Red
    Write-Host "Please install Python 3.11 first:" -ForegroundColor Yellow
    Write-Host "1. Download from https://www.python.org/downloads/release/python-3119/" -ForegroundColor White
    Write-Host "2. Install with 'Add Python to PATH' checked" -ForegroundColor White
    Write-Host "3. Run this script again" -ForegroundColor White
    Read-Host "`nPress Enter to exit"
    exit 1
}

Write-Host "Using: $pythonExe" -ForegroundColor Green
& $pythonExe --version
Write-Host ""

# Navigate to project
$projectPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectPath
Write-Host "Project: $projectPath" -ForegroundColor Green

# Create necessary directories
$directories = @("output", "logs", "cache", "backups", "config")
foreach ($dir in $directories) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
        Write-Host "Created directory: $dir" -ForegroundColor Yellow
    }
}

# Check and create default configuration
$configFile = "config/dashboard_config.yaml"
if (-not (Test-Path $configFile)) {
    $exampleConfig = "config/dashboard_config.yaml.example"
    if (Test-Path $exampleConfig) {
        Copy-Item $exampleConfig $configFile
        Write-Host "Created default configuration" -ForegroundColor Yellow
    }
    else {
        Write-Host "Creating default configuration file..." -ForegroundColor Yellow
        # Create basic config
        $configContent = @"
dashboard_config:
  social_media:
    keywords:
      - "TF Value-Mart FB Page Wallposts Performance"
      - "TF Value-Mart IG Page Wallposts Performance"
      - "TF Value-Mart Tiktok Video Performance"
    platforms:
      Facebook:
        metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
      Instagram:
        metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
      TikTok:
        metrics: ["Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
"@
        Set-Content -Path $configFile -Value $configContent
    }
}

# Check dependencies
Write-Host "`nChecking dependencies..." -ForegroundColor Yellow
try {
    & $pythonExe -c "import pandas, pptx, plotly, yaml, numpy" 2>$null
    Write-Host "✓ All dependencies installed" -ForegroundColor Green
} catch {
    Write-Host "Installing required packages..." -ForegroundColor Yellow
    & $pythonExe -m pip install --upgrade pip
    & $pythonExe -m pip install python-pptx pandas openpyxl plotly pyyaml numpy
}

Write-Host ""
Write-Host "Starting application..." -ForegroundColor Green
Write-Host "Press Ctrl+C to stop" -ForegroundColor Gray
Write-Host ""

# Run the application
try {
    & $pythonExe src/main.py
} catch {
    Write-Host "`nApplication error: $_" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"