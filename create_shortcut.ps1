$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Output to Windows Startup folder so kotonoha launches at login
# and Whisper/Ollama warm up in the background before first use.
$startup = [Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startup 'VoiceInput.lnk'

$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) {
    Write-Host 'pythonw.exe not found in PATH'
    exit 1
}

$projectDir = 'C:\ClaudeCode\projects\kotonoha'
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

Write-Host "Created: $shortcutPath"
Write-Host "Target : $pythonw"
Write-Host "WorkDir: $projectDir"
Write-Host "Exists : $(Test-Path $shortcutPath)"
Write-Host ''
Write-Host 'Registered to Windows Startup. It will auto-launch at next login.'
Write-Host 'To remove: delete the .lnk from this folder:'
Write-Host "  $startup"
