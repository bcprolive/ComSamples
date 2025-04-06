<#
.SYNOPSIS
  COM Samples Setup Script with Robust Error Handling

.DESCRIPTION
  Final version with:
  - Validated path construction
  - PowerShell 5.1+ compatibility
  - Comprehensive environment validation
  - Detailed error diagnostics
#>

# --- Step 0: Initialize Logging ---
Start-Transcript -Path "$PSScriptRoot\SetupComSamples.log" -Append
Write-Host "Starting ComSamples setup at $(Get-Date)" -ForegroundColor Cyan

# --- Step 1: Administrator Check ---
function Test-Administrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script must be run as Administrator. Exiting."
        exit 1
    }
}
Test-Administrator

# --- Step 2: Preflight Checks ---
$requiredTools = @('git', 'python')
foreach ($tool in $requiredTools) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Write-Error "$tool is not installed or not in PATH."
        exit 1
    }
}

# --- Step 3: Visual Studio Validation ---
$vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswherePath)) {
    Write-Error "Visual Studio Installer components missing. Install Visual Studio 2022+."
    exit 1
}

# --- Step 4: ATL Component Verification ---
$atlComponents = @(
    "Microsoft.VisualStudio.Component.VC.ATL",
    "Microsoft.VisualStudio.Component.VC.ATLMFC"
)

foreach ($component in $atlComponents) {
    $componentCheck = & $vswherePath -latest -requires $component
    if (-not $componentCheck) {
        Write-Error "Missing required component: $component"
        Write-Host "Install via Visual Studio Installer > Modify > Individual Components > Search 'ATL'"
        exit 1
    }
}

# --- Step 5: MSBuild Detection ---
$vsInstallPath = & $vswherePath -latest -property installationPath
$msbuildPath = & $vswherePath -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe"

if (-not $msbuildPath) {
    Write-Error "MSBuild not found. Install Visual Studio with C++ workload."
    exit 1
}
Write-Host "MSBuild located at: $msbuildPath"

# --- Step 6: ATL Version Detection (Fixed) ---
$msvcDir = Join-Path $vsInstallPath "VC\Tools\MSVC"
$atlVersions = Get-ChildItem -Path $msvcDir -Directory | Where-Object {
    Test-Path (Join-Path $_.FullName "atlmfc\include\atlbase.h") -PathType Leaf
} | Sort-Object {
    if ($_.Name -match '(\d+)\.(\d+)\.(\d+)') {
        [tuple]::Create([int]$matches[1], [int]$matches[2], [int]$matches[3])
    }
} -Descending

if (-not $atlVersions) {
    Write-Error "No valid ATL versions detected. Components missing: C++ ATL for v143 build tools"
    exit 1
}

$selectedAtlVersion = $atlVersions[0].FullName
Write-Host "Selected ATL Version: $selectedAtlVersion"

# --- Step 7: Null-Safe Path Construction ---
$atlIncludePath = Join-Path $selectedAtlVersion "atlmfc\include" -ErrorAction Stop
$vcIncludePath = Join-Path $selectedAtlVersion "include" -ErrorAction Stop

if (-not (Test-Path $atlIncludePath)) {
    Throw "Critical ATL include path missing: $atlIncludePath"
}

# --- Step 8: INCLUDE Path Configuration (Validated) ---
$includePaths = @(
    $atlIncludePath,
    $vcIncludePath
) | Where-Object { Test-Path $_ }

if ($includePaths.Count -eq 0) {
    Write-Error "Failed to construct valid INCLUDE paths. Verify Visual Studio installation."
    exit 1
}

$env:INCLUDE = ($includePaths + ($env:INCLUDE -split ';' | Where-Object { $_ })) -join ';'
Write-Host "Configured INCLUDE paths:`n$($env:INCLUDE -replace ';', "`n")"
# --- Remaining sections remain unchanged ---

# --- Step 9: Environment Configuration (Fixed) ---
$vcvarsall = Join-Path $vsInstallPath "VC\Auxiliary\Build\vcvarsall.bat"
$envFile = Join-Path $env:TEMP "vcenv_$(Get-Date -Format 'yyyyMMddHHmmss').txt"

# Create temporary batch file
@"
@echo off
call "{0}" x64 -vcvars_ver={1}
set > "{2}"
"@ -f $vcvarsall, $vcToolsVersion, $envFile | Set-Content "$envFile.bat" -Encoding ASCII

# Execute and capture environment
$process = Start-Process cmd.exe -ArgumentList "/c `"$envFile.bat`"" `
    -Wait -PassThru -WindowStyle Hidden

if ($process.ExitCode -ne 0) {
    Write-Error "Failed to configure Visual Studio environment"
    exit 1
}

# Load environment variables
Get-Content $envFile | ForEach-Object {
    if ($_ -match '^([^=]+)=(.*)$') {
        [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2], 'Process')
    }
}

# Cleanup
Remove-Item "$envFile.bat", $envFile -ErrorAction SilentlyContinue

# --- Step 10: INCLUDE Path Validation ---
$requiredIncludes = @(
    (Join-Path $selectedAtlVersion "atlmfc\include"),
    (Join-Path $selectedAtlVersion "include"),
    (Join-Path $env:WindowsSdkDir "Include\$($env:WindowsSDKVersion)\um"),
    (Join-Path $env:WindowsSdkDir "Include\$($env:WindowsSDKVersion)\shared")
)

$validIncludes = $requiredIncludes | Where-Object { Test-Path $_ }
if ($validIncludes.Count -lt 4) {
    Write-Error "Missing critical include paths:"
    $requiredIncludes | ForEach-Object { 
        if (-not (Test-Path $_)) { Write-Host "  Missing: $_" }
    }
    exit 1
}

$env:INCLUDE = ($validIncludes + ($env:INCLUDE -split ';' | Where-Object { $_ })) -join ';'

# --- Step 11: MSBuild Execution (Fixed) ---
$buildParams = @(
    "ComSamples.sln",
    "/p:Configuration=Release",
    "/p:Platform=x64",
    "/p:PreferredToolArchitecture=x64",
    "/v:minimal"
)

Write-Host "Building with verified environment:" -ForegroundColor Cyan
Write-Host "INCLUDE: $($env:INCLUDE -replace ';', "`n  ")"
Write-Host "LIB: $($env:LIB -replace ';', "`n  ")"

& "$msbuildPath" $buildParams
if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed with exit code $LASTEXITCODE"
    exit 1
}

# --- Step 12: Fix MIDL-Generated File Handling ---
$projects = @(
    "MyInterfaces\MyInterfaces.vcxproj",
    "ServiceWrapper\ServiceWrapper.vcxproj"
)

foreach ($proj in $projects) {
    $projPath = Join-Path $PSScriptRoot $proj
    $projContent = Get-Content $projPath -Raw

    # Force C++ compilation for all files
    $projContent = $projContent -replace 
        '<ItemDefinitionGroup>',
        '<ItemDefinitionGroup><ClCompile><CompileAs>CompileAsCpp</CompileAs></ClCompile>'

    Set-Content -Path $projPath -Value $projContent -Force
}

# --- Step 13: Rebuild Solution ---
Write-Host "Rebuilding with C++ compilation enforced..." -ForegroundColor Cyan
& $msbuildPath ComSamples.sln /p:Configuration=Release /p:Platform=x64 /v:minimal /t:Clean,Build

# --- Final Output ---
Write-Host "`nSetup completed. Client applications:" -ForegroundColor Green
Write-Host "  C++: MyClientCpp\bin\x64\Release\MyClientCpp.exe"
Write-Host "  C#:  MyClientCs\bin\Release\MyClientCs.exe"
Write-Host "  Python: python MyClientPy\bin\Release\MyClientPy.py"

Stop-Transcript
