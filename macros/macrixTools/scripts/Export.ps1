param(
    [string]$OutputRoot = (Join-Path $PSScriptRoot "..\export")
)

New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
Write-Host "Export placeholder. Use Build.ps1 for source packaging and Install.ps1 for CorelDRAW deployment."
