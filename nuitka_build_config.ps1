# Universal Nuitka folder build script.
# Put this file next to the main .py file, edit Settings, then run:
#   cd "D:\DIR\Temporal_Files\Датчик давления"
#   powershell -ExecutionPolicy Bypass -File .\nuitka_build_config.ps1
# =========================
# Settings
# =========================

# Main Python file to compile. Relative paths are resolved from this .ps1 folder.
$MainScript = "rfid_access_qt_logger_journal_only.py"

# Python launcher to use for dependency checks and Nuitka.
$PythonCommand = "py"

# Final output folder name. It will be created next to this .ps1 file.
$FinalFolderName = "Access Control System"

# Final executable name inside the output folder.
$FinalExeName = "Access Control System.exe"

# Main executable icon. Leave empty to skip.
# Example: "assets/app.ico" or "D:\icons\app.ico"
$AppIconPath = "KeyChainAccess_37052.ico"

# Extra icon file to copy into the output folder for your app to load at runtime.
# Leave empty to skip. Your Python code still needs to load this file with QIcon.
$PanelIconPath = "KeyChainAccess_37052.ico"
$PanelIconTarget = "KeyChainAccess_37052.ico"

# Extra data files to include. Format: source=target.
# Relative source paths are resolved from this .ps1 folder.
$DataFiles = @(
    ""
)

# Nuitka plugins to enable. Remove plugins that your project does not use.
$NuitkaPlugins = @(
    "pyqt6"
)

# Python packages that must be bundled explicitly.
$IncludedPackages = @(
    "PyQt6",
    "serial",
    "openpyxl"
)

# Imports checked before build so missing dependencies fail with a clear message.
$RequiredImports = @(
    @{ ImportName = "PyQt6"; PackageName = "PyQt6" },
    @{ ImportName = "serial"; PackageName = "pyserial" },
    @{ ImportName = "openpyxl"; PackageName = "openpyxl" }
)

# Extra Nuitka arguments, one item per argument.
$ExtraNuitkaArgs = @(
    "--windows-console-mode=disable"
)

# =========================
# Build logic
# =========================

$ProjectDir = $PSScriptRoot
$MainScriptPath = Join-Path $ProjectDir $MainScript
$MainScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($MainScript)
$NuitkaDistPath = Join-Path $ProjectDir "$MainScriptBaseName.dist"
$FinalFolderPath = Join-Path $ProjectDir $FinalFolderName
$OriginalExePath = Join-Path $FinalFolderPath "$MainScriptBaseName.exe"
$FinalExePath = Join-Path $FinalFolderPath $FinalExeName

function Resolve-ProjectPath {
    param([string] $Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path $ProjectDir $Path
}

function Resolve-DataFileArg {
    param([string] $DataFile)

    $parts = $DataFile -split "=", 2
    if ($parts.Count -ne 2) {
        throw "Bad data file entry '$DataFile'. Use source=target."
    }

    $source = Resolve-ProjectPath $parts[0]
    $target = $parts[1]
    return "$source=$target"
}

if (-not (Test-Path -LiteralPath $MainScriptPath)) {
    throw "Main script not found: $MainScriptPath"
}

foreach ($requiredImport in $RequiredImports) {
    & $PythonCommand -c "import $($requiredImport.ImportName)" 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "Missing Python package '$($requiredImport.PackageName)'. Install it with: py -m pip install $($requiredImport.PackageName)"
    }
}

$NuitkaArgs = @(
    "--standalone",
    "--output-dir=$ProjectDir"
)

foreach ($arg in $ExtraNuitkaArgs) {
    if (-not [string]::IsNullOrWhiteSpace($arg)) {
        $NuitkaArgs += $arg
    }
}

foreach ($plugin in $NuitkaPlugins) {
    if (-not [string]::IsNullOrWhiteSpace($plugin)) {
        $NuitkaArgs += "--enable-plugin=$plugin"
    }
}

foreach ($package in $IncludedPackages) {
    if (-not [string]::IsNullOrWhiteSpace($package)) {
        $NuitkaArgs += "--include-package=$package"
    }
}

$ResolvedAppIconPath = Resolve-ProjectPath $AppIconPath
if ($ResolvedAppIconPath) {
    if (-not (Test-Path -LiteralPath $ResolvedAppIconPath)) {
        throw "App icon not found: $ResolvedAppIconPath"
    }
    $NuitkaArgs += "--windows-icon-from-ico=$ResolvedAppIconPath"
}

$ResolvedPanelIconPath = Resolve-ProjectPath $PanelIconPath
if ($ResolvedPanelIconPath) {
    if (-not (Test-Path -LiteralPath $ResolvedPanelIconPath)) {
        throw "Panel icon not found: $ResolvedPanelIconPath"
    }
    $NuitkaArgs += "--include-data-files=$ResolvedPanelIconPath=$PanelIconTarget"
}

foreach ($dataFile in $DataFiles) {
    if (-not [string]::IsNullOrWhiteSpace($dataFile)) {
        $NuitkaArgs += "--include-data-files=$(Resolve-DataFileArg $dataFile)"
    }
}

$NuitkaArgs += $MainScriptPath

& $PythonCommand -m nuitka @NuitkaArgs

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

if (Test-Path -LiteralPath $FinalFolderPath) {
    Remove-Item -LiteralPath $FinalFolderPath -Recurse -Force
}

Move-Item -LiteralPath $NuitkaDistPath -Destination $FinalFolderPath

if ((Test-Path -LiteralPath $OriginalExePath) -and ($OriginalExePath -ne $FinalExePath)) {
    Move-Item -LiteralPath $OriginalExePath -Destination $FinalExePath -Force
}

Write-Output "Done: $FinalExePath"

# Удаление папки .build после сборки
$buildDirs = Get-ChildItem -Path . -Directory -Filter "*.build" -Recurse

foreach ($dir in $buildDirs) {
    try {
        Remove-Item -Recurse -Force -Path $dir.FullName
        Write-Host "Удалена папка: $($dir.FullName)"
    } catch {
        Write-Host "Не удалось удалить: $($dir.FullName)"
    }
}