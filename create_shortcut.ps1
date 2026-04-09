$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'VoiceInput.lnk'

$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) {
    Write-Host 'pythonw.exe not found in PATH'
    exit 1
}

$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut($shortcutPath)
$s.TargetPath = $pythonw
$s.Arguments = 'voice_input.py'
$s.WorkingDirectory = 'C:\ClaudeCode\tools\voice-input'
$s.IconLocation = 'C:\ClaudeCode\tools\voice-input\mic.ico'
$s.Description = 'Voice Input Tool (Whisper + Qwen)'
$s.WindowStyle = 7  # Minimized
$s.Save()

Write-Host "Created: $shortcutPath"
Write-Host "Target : $pythonw"
Write-Host "Exists : $(Test-Path $shortcutPath)"
