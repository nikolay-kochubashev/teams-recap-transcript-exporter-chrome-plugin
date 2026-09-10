$ErrorActionPreference = 'Stop'

$HostName = 'com.openai.teams_recap_transcript_exporter'
$ExtensionId = 'nlnkhodhdodapoogefbfjgejodgnfmba'
$InstallDir = Join-Path $env:LOCALAPPDATA 'TeamsRecapTranscriptExporter'
$ExePath = Join-Path $InstallDir 'TeamsRecapNativeHost.exe'
$ManifestPath = Join-Path $InstallDir ($HostName + '.json')
$SourcePath = Join-Path $PSScriptRoot 'NativeHost.cs'

if (-not (Test-Path $SourcePath)) { throw "NativeHost.cs not found: $SourcePath" }
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
if (Test-Path $ExePath) { Remove-Item -Force $ExePath }

Add-Type -Path $SourcePath -ReferencedAssemblies 'System.Web.Extensions.dll' -OutputAssembly $ExePath -OutputType ConsoleApplication

$manifest = [ordered]@{
    name = $HostName
    description = 'Teams Recap Transcript Exporter Windows integration'
    path = $ExePath
    type = 'stdio'
    allowed_origins = @("chrome-extension://$ExtensionId/")
} | ConvertTo-Json -Depth 4

[System.IO.File]::WriteAllText($ManifestPath, $manifest, (New-Object System.Text.UTF8Encoding($false)))

$RegistryPath = "HKCU:\Software\Google\Chrome\NativeMessagingHosts\$HostName"
New-Item -Path $RegistryPath -Force | Out-Null
Set-Item -Path $RegistryPath -Value $ManifestPath

Write-Host ''
Write-Host 'Windows integration installed successfully.' -ForegroundColor Green
Write-Host "Host: $ExePath"
Write-Host "Documents: $([Environment]::GetFolderPath('MyDocuments'))"
Write-Host ''
Write-Host 'Now reload the extension in chrome://extensions and reopen its side panel.'
