param(
    [switch]$Startup
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Default: create a shortcut on the Desktop.
# With -Startup: place the shortcut in the Windows Startup folder so
# kotonoha auto-launches at login and Whisper/Ollama warm up in the
# background before first use.
if ($Startup) {
    $targetDir = [Environment]::GetFolderPath('Startup')
    $mode = 'Startup'
} else {
    $targetDir = [Environment]::GetFolderPath('Desktop')
    $mode = 'Desktop'
}
$shortcutPath = Join-Path $targetDir 'VoiceInput.lnk'

$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) {
    Write-Host 'pythonw.exe not found in PATH'
    exit 1
}

$projectDir = $PSScriptRoot
if (-not $projectDir) {
    $projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$iconPath = Join-Path $projectDir 'mic.ico'

if (-not (Test-Path (Join-Path $projectDir 'voice_input.py'))) {
    Write-Host "voice_input.py not found under: $projectDir"
    exit 1
}

$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut($shortcutPath)
$s.TargetPath = $pythonw
$s.Arguments = 'voice_input.py'
$s.WorkingDirectory = $projectDir
$s.IconLocation = $iconPath
$s.Description = 'Voice Input Tool (Whisper + Qwen)'
$s.WindowStyle = 7  # Minimized
$s.Save()

Write-Host "Mode    : $mode"
Write-Host "Created : $shortcutPath"
Write-Host "Target  : $pythonw"
Write-Host "WorkDir : $projectDir"
Write-Host "Exists  : $(Test-Path $shortcutPath)"
Write-Host ''
if ($Startup) {
    Write-Host 'Registered to Windows Startup. It will auto-launch at next login.'
    Write-Host 'To remove: delete the .lnk from this folder:'
    Write-Host "  $targetDir"
} else {
    Write-Host 'Created on Desktop. To auto-launch at login instead, re-run with:'
    Write-Host '  .\create_shortcut.ps1 -Startup'
}
