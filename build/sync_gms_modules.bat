@echo off
rem ==============================================================================
rem sync_gms_modules.bat
rem ------------------------------------------------------------------------------
rem Copies VBA module files (.bas and .frm) from the repository's src\ directory
rem into the CorelDRAW GMS folder where Macrix.gms resides, staging them for
rem import into the live Macrix.gms project via the CorelDRAW VBA editor.
rem
rem Usage:
rem   sync_gms_modules.bat [gms_folder_path]
rem
rem   If no argument is given the script auto-detects the CorelDRAW GMS folder.
rem   You can also override detection by setting the MACRIX_GMS_DIR environment
rem   variable.  A command-line argument always takes precedence.
rem
rem Exit codes:
rem   0  success (files copied)
rem   1  source directory not found
rem   2  GMS destination directory not found / not writable
rem   3  no .bas/.frm files found in source
rem ==============================================================================
setlocal EnableDelayedExpansion

rem --- Source directory: the repository src\ one level up from build\ ---
set "SRC_DIR=%~dp0..\src"

rem --- Destination: CLI arg takes precedence, then env var, then auto-detect ---
set "GMS_DIR="
if not "%~1"=="" (
    set "GMS_DIR=%~1"
)
if "%GMS_DIR%"=="" (
    if defined MACRIX_GMS_DIR set "GMS_DIR=%MACRIX_GMS_DIR%"
)

rem --- Auto-detect the CorelDRAW GMS folder if still unset ---
if "%GMS_DIR%"=="" (
    call :DetectGMSDir
)

rem --- Validate source ---
if not exist "%SRC_DIR%\" (
    echo ERROR: Source directory not found: %SRC_DIR%
    exit /b 1
)

rem --- Validate / create destination ---
if "%GMS_DIR%"=="" (
    echo ERROR: Could not auto-detect the CorelDRAW GMS folder.
    echo.
    echo CorelDRAW stores GMS files under its user data directory, e.g.:
    echo   %%APPDATA%%\Corel\CorelDRAW Graphics Suite 2021\Draw\GMS
    echo   %%APPDATA%%\Corel\CorelDRAW Graphics Suite 2020\Draw\GMS
    echo   %%APPDATA%%\Corel\CorelDRAW X8\18.0\GMS
    echo.
    echo Fix this by either:
    echo   1. Passing the GMS path as an argument:  %~nx0 "C:\path\to\GMS"
    echo   2. Setting the MACRIX_GMS_DIR environment variable to the GMS path.
    exit /b 2
)

if not exist "%GMS_DIR%\" (
    echo ERROR: Destination GMS directory does not exist: %GMS_DIR%
    exit /b 2
)

rem --- Count source files ---
set "BAS_COUNT=0"
set "FRM_COUNT=0"
for %%F in ("%SRC_DIR%\*.bas") do set /a BAS_COUNT+=1
for %%F in ("%SRC_DIR%\*.frm") do set /a FRM_COUNT+=1
set /a TOTAL_COUNT=BAS_COUNT+FRM_COUNT

if %TOTAL_COUNT%==0 (
    echo ERROR: No .bas or .frm files found in %SRC_DIR%
    exit /b 3
)

echo.
echo Copying VBA modules to CorelDRAW GMS folder
echo.
echo   Source : %SRC_DIR%
echo   Target : %GMS_DIR%
echo   Files  : %TOTAL_COUNT% ^( %BAS_COUNT% .bas, %FRM_COUNT% .frm ^)
echo.

rem --- Copy .bas files ---
set "COPY_OK=0"
set "COPY_FAIL=0"
for %%F in ("%SRC_DIR%\*.bas") do (
    copy /Y "%%F" "%GMS_DIR%" >nul 2>&1
    if !errorlevel!==0 (
        echo   [OK]   %%~nxF
        set /a COPY_OK+=1
    ) else (
        echo   [FAIL] %%~nxF
        set /a COPY_FAIL+=1
    )
)

rem --- Copy .frm files ---
for %%F in ("%SRC_DIR%\*.frm") do (
    copy /Y "%%F" "%GMS_DIR%" >nul 2>&1
    if !errorlevel!==0 (
        echo   [OK]   %%~nxF
        set /a COPY_OK+=1
    ) else (
        echo   [FAIL] %%~nxF
        set /a COPY_FAIL+=1
    )
)

echo.
echo Copied !COPY_OK! of !TOTAL_COUNT! file^(s^).
if !COPY_FAIL! gtr 0 (
    echo WARNING: %COPY_FAIL% file^(s^) failed to copy.
)

rem --- Check if Macrix.gms exists in the target folder ---
if exist "%GMS_DIR%\macrix.gms" (
    echo.
    echo NOTE: macrix.gms was found in the GMS folder.
    echo   To import these modules into the live project, open CorelDRAW,
    echo   open the VBA editor ^(Alt+F11^), right-click the Macrix macro project,
    echo   and use File ^> Import File... for each .bas file listed above.
)

if !COPY_FAIL! gtr 0 (
    exit /b 2
)

exit /b 0

rem ==============================================================================
rem DetectGMSDir - locate the CorelDRAW GMS folder by searching common paths.
rem Sets the GMS_DIR variable on success.
rem ==============================================================================
:DetectGMSDir
rem Try %APPDATA%\Corel (Roaming - most common for CorelDRAW)
set "COREL_ROOT=%APPDATA%\Corel"
if not exist "%COREL_ROOT%" (
    rem Fall back to %LOCALAPPDATA%\Corel
    set "COREL_ROOT=%LOCALAPPDATA%\Corel"
)
if not exist "%COREL_ROOT%" goto :eof

rem Pattern 1: CorelDRAW Graphics Suite YYYY\Draw\GMS  (2019, 2020, 2021+)
for /d %%A in ("%COREL_ROOT%\CorelDRAW Graphics Suite *") do (
    if exist "%%A\Draw\GMS\" (
        set "GMS_DIR=%%A\Draw\GMS"
        goto :eof
    )
    if exist "%%A\GMS\" (
        set "GMS_DIR=%%A\GMS"
        goto :eof
    )
    rem Some older sub-versions: ...\Suite YYYY\version\GMS
    for /d %%B in ("%%A\1*") do (
        if exist "%%B\GMS\" (
            set "GMS_DIR=%%B\GMS"
            goto :eof
        )
    )
)

rem Pattern 2: CorelDRAW X*  (X5, X6, X7, X8 - legacy naming)
for /d %%A in ("%COREL_ROOT%\CorelDRAW *") do (
    if exist "%%A\Draw\GMS\" (
        set "GMS_DIR=%%A\Draw\GMS"
        goto :eof
    )
    rem CorelDRAW X8 structure: CorelDRAW X8\18.0\GMS
    for /d %%B in ("%%A\*") do (
        if exist "%%B\GMS\" (
            set "GMS_DIR=%%B\GMS"
            goto :eof
        )
    )
    if exist "%%A\GMS\" (
        set "GMS_DIR=%%A\GMS"
        goto :eof
    )
)

rem Pattern 3: generic search for any GMS folder under CorelDRAW product dirs
for /d %%A in ("%COREL_ROOT%\CorelDRAW *") do (
    for /d %%B in ("%%A\*") do (
        for /d %%C in ("%%B\*") do (
            if exist "%%C\GMS\" (
                set "GMS_DIR=%%C\GMS"
                goto :eof
            )
        )
        if exist "%%B\GMS\" (
            set "GMS_DIR=%%B\GMS"
            goto :eof
        )
    )
)
goto :eof
