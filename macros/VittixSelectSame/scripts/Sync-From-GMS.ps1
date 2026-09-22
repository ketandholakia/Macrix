param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot ".."),
    [string]$ProjectName = "VittixSelectSame"
)

$srcPath = Join-Path $SourceRoot "src"
$formsPath = Join-Path $SourceRoot "forms"

Write-Host "Connecting to CorelDRAW..."
try {
    $cdr = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application.23")
} catch {
    try {
        $cdr = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application")
    } catch {
        Write-Host "Could not find a running CorelDRAW instance. Starting a background instance..."
        $cdr = New-Object -ComObject CorelDRAW.Application
        Start-Sleep -Seconds 2
    }
}

try {
    $vbe = $cdr.VBE
} catch {
    Write-Error "Failed to access the VBA Extensibility model. Please ensure your Macro Security settings allow macro execution."
    exit
}

$project = $vbe.VBProjects | Where-Object { $_.Name -eq $ProjectName }

if ($null -eq $project) {
    Write-Error "Could not find a loaded Global Macro project named '$ProjectName'."
    exit
}

Write-Host "Found project: $($project.Name)"

foreach ($comp in $project.VBComponents) {
    # Component Types: 1=Standard Module, 2=Class Module, 3=UserForm, 100=Document
    if ($comp.Type -eq 1) {
        $exportPath = Join-Path $srcPath "$($comp.Name).bas"
        Write-Host "Exporting Module: $($comp.Name) -> $exportPath"
        $comp.Export($exportPath)
    } elseif ($comp.Type -eq 2) {
        $exportPath = Join-Path $srcPath "$($comp.Name).cls"
        Write-Host "Exporting Class: $($comp.Name) -> $exportPath"
        $comp.Export($exportPath)
    } elseif ($comp.Type -eq 3) {
        $exportPath = Join-Path $formsPath "$($comp.Name).frm"
        Write-Host "Exporting Form: $($comp.Name) -> $exportPath"
        $comp.Export($exportPath)
        # Note: Exporting the .frm automatically exports the .frx next to it
    }
}

Write-Host "Successfully exported code from CorelDRAW GMS project to source directory."
