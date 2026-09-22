param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot ".."),
    [string]$ProjectName = "vittixBleed"
)

$srcPath = Join-Path $SourceRoot "src"
$formsPath = Join-Path $SourceRoot "forms"

Write-Host "Connecting to CorelDRAW..."
try {
    # Try to connect to an existing running instance of CorelDRAW 
    $cdr = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application")
} catch {
    # If not running, start it
    Write-Host "CorelDRAW is not running. Starting a new instance..."
    $cdr = New-Object -ComObject CorelDRAW.Application
}

# Ensure CorelDRAW is fully loaded
Start-Sleep -Seconds 2

try {
    $vbe = $cdr.VBE
} catch {
    Write-Error "Failed to access the VBA Extensibility model. Please ensure your Macro Security settings allow macro execution."
    exit
}

$project = $vbe.VBProjects | Where-Object { $_.Name -eq $ProjectName }

if ($null -eq $project) {
    Write-Error "Could not find a loaded Global Macro project named '$ProjectName'. Please create an empty macro project in CorelDRAW with this name and save it first."
    exit
}

Write-Host "Found project: $($project.Name)"

# Remove existing components (except for ThisMacroStorage)
$componentsToRemove = $project.VBComponents | Where-Object { $_.Type -ne 100 } # 100 is typically Document/ThisDocument
foreach ($comp in $componentsToRemove) {
    Write-Host "Removing old component: $($comp.Name)"
    $project.VBComponents.Remove($comp)
}

# Import BAS files
Write-Host "Importing modules from $srcPath..."
Get-ChildItem -Path $srcPath -Filter "*.bas" | ForEach-Object {
    Write-Host "  -> $($_.Name)"
    $comp = $project.VBComponents.Import($_.FullName)
    if ($comp.Name -ne $_.BaseName) {
        try {
            $comp.Name = $_.BaseName
        } catch {
            Write-Host "  -> Warning: Could not rename to $($_.BaseName). It remains named $($comp.Name)" -ForegroundColor Yellow
        }
    }
}

# Import CLS files
Write-Host "Importing class modules from $srcPath..."
Get-ChildItem -Path $srcPath -Filter "*.cls" | ForEach-Object {
    Write-Host "  -> $($_.Name)"
    if ($_.BaseName -eq "ThisMacroStorage") {
        Write-Host "  -> Syncing ThisMacroStorage code..."
        $docComp = $project.VBComponents.Item("ThisMacroStorage")
        $docCode = $docComp.CodeModule
        if ($docCode.CountOfLines -gt 0) {
            $docCode.DeleteLines(1, $docCode.CountOfLines)
        }
        
        # Read the file and strip the VBA class headers manually
        $lines = Get-Content $_.FullName
        $codeLines = @()
        $inHeader = $true
        foreach ($line in $lines) {
            if ($inHeader) {
                if ($line -notmatch "^VERSION " -and $line -notmatch "^BEGIN" -and $line -notmatch "^END" -and $line -notmatch "^Attribute " -and $line -notmatch "^\s*MultiUse\s*=") {
                    $inHeader = $false
                }
            }
            if (-not $inHeader) {
                $codeLines += $line
            }
        }
        
        $codeText = $codeLines -join "`r`n"
        if (-not [string]::IsNullOrWhiteSpace($codeText)) {
            $docCode.AddFromString($codeText)
        }
    } else {
        $comp = $project.VBComponents.Import($_.FullName)
        if ($comp.Name -ne $_.BaseName) {
            try {
                $comp.Name = $_.BaseName
            } catch {
                Write-Host "  -> Warning: Could not rename to $($_.BaseName). It remains named $($comp.Name)" -ForegroundColor Yellow
            }
        }
    }
}

# Import FRM files (which also auto-imports FRX if it exists)
Write-Host "Importing forms from $formsPath..."
Get-ChildItem -Path $formsPath -Filter "*.frm" | ForEach-Object {
    Write-Host "  -> $($_.Name)"
    $comp = $project.VBComponents.Import($_.FullName)
    # Forms usually retain their name from internal attributes, but we can enforce it just in case:
    $comp.Name = $_.BaseName
}

Write-Host "Successfully synced source code into the CorelDRAW GMS project."
Write-Host "Remember to press 'Save' inside the CorelDRAW Macro Editor to persist changes to the actual .gms file."
