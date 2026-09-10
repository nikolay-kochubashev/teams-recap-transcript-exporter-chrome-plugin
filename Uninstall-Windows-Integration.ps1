$ErrorActionPreference = 'Stop'
$HostName = 'com.openai.teams_recap_transcript_exporter'
$RegistryPath = "HKCU:\Software\Google\Chrome\NativeMessagingHosts\$HostName"
$InstallDir = Join-Path $env:LOCALAPPDATA 'TeamsRecapTranscriptExporter'
if (Test-Path $RegistryPath) { Remove-Item -Path $RegistryPath -Recurse -Force }
if (Test-Path $InstallDir) { Remove-Item -Path $InstallDir -Recurse -Force }
Write-Host 'Windows integration removed.' -ForegroundColor Green
