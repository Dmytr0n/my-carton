@echo off
chcp 65001 >nul
setlocal

REM =========================================================
REM MY_CARTON BUILD SCRIPT
REM =========================================================

set "rootDir=%cd%"

REM =========================================================
REM STATUS VARIABLES
REM =========================================================

set "step1Status=NOT STARTED"
set "step2Status=NOT STARTED"
set "step3Status=NOT STARTED"
set "step4Status=NOT STARTED"
set "step5Status=NOT STARTED"
set "step8Status=NOT STARTED"
set "step9Status=NOT STARTED"

REM =========================================================
REM PATHS
REM =========================================================

set "solutionPath=%rootDir%\src\MyCarton\moy_carton.sln"
set "clientBuildOutput=%rootDir%\deploy\client"
set "artifactDir=%rootDir%\artefacts"
set "clientArtifactZipPath=%artifactDir%\my_carton_build.zip"

REM =========================================================
REM STEP 1 - CREATE FOLDERS
REM =========================================================

echo.
echo =========================================================
echo Step 1: Setting up folder structure...
echo =========================================================

if not exist "%rootDir%\deploy" (
    mkdir "%rootDir%\deploy"
)

if not exist "%clientBuildOutput%" (
    mkdir "%clientBuildOutput%"
)

if not exist "%artifactDir%" (
    mkdir "%artifactDir%"
)

set "step1Status=PASSED"

REM =========================================================
REM START BUILD
REM =========================================================

echo.
echo =========================================================
echo MOY_KARTON BUILD START
echo =========================================================

REM =========================================================
REM STEP 2 - CHECK SOLUTION
REM =========================================================

echo.
echo Step 2: Checking solution file...

if exist "%solutionPath%" (
    echo [OK] Solution found.
    set "step2Status=PASSED"
) else (
    echo [ERROR] Solution file not found:
    echo %solutionPath%
    set "step2Status=FAILED"
    goto FinalReport
)

REM =========================================================
REM STEP 3 - RESTORE NUGET
REM =========================================================

echo.
echo Step 3: Restoring NuGet packages...

nuget restore "%solutionPath%"

if not "%errorlevel%"=="0" (
    echo [ERROR] Failed to restore NuGet packages.
    set "step3Status=FAILED"
    goto FinalReport
)

echo [OK] NuGet packages restored successfully.
set "step3Status=PASSED"

REM =========================================================
REM STEP 4 - CHECK MSBUILD
REM =========================================================

echo.
echo Step 4: Checking for MSBuild...

where msbuild >nul 2>nul

if errorlevel 1 (
    echo [ERROR] MSBuild.exe not found.
    set "step4Status=FAILED"
    goto FinalReport
)

echo [OK] MSBuild found.
set "step4Status=PASSED"

REM =========================================================
REM STEP 5 - BUILD PROJECT
REM =========================================================

echo.
echo Step 5: Building project...

msbuild "%solutionPath%" ^
/p:Configuration=Release ^
/p:OutputPath="%clientBuildOutput%" ^
/p:Platform="Any CPU"

if exist "%clientBuildOutput%\moy_carton.exe" (
    echo [OK] Build completed successfully.
    set "step5Status=PASSED"
) else (
    echo [ERROR] Build failed.
    set "step5Status=FAILED"
    goto FinalReport
)

REM =========================================================
REM ВСТАВИТИ КОД ДЛЯ КОПІЮВАННЯ ІКОНКИ ТУТ
REM =========================================================
echo.
echo Step 5.1: Copying icon to deploy folder...

copy /Y "%rootDir%\src\MyCarton\karton.ico" "%clientBuildOutput%\karton.ico" >nul

if "%errorlevel%"=="0" (
    echo [OK] karton.ico copied successfully.
) else (
    echo [WARNING] Failed to copy karton.ico.
)

REM =========================================================
REM =========================================================
REM STEP 8 - CREATE ARCHIVE
REM =========================================================

echo.
echo Step 8: Creating build artifact archive...

if exist "%clientArtifactZipPath%" (
    del /f /q "%clientArtifactZipPath%"
)

powershell -Command "Compress-Archive -Path '%clientBuildOutput%\*' -DestinationPath '%clientArtifactZipPath%' -Force"

if not "%errorlevel%"=="0" (
    echo [ERROR] Failed to create archive.
    set "step8Status=FAILED"
) else (
    echo [OK] Artifact saved at:
    echo %clientArtifactZipPath%
    set "step8Status=PASSED"
)

REM =========================================================
REM STEP 9 - BUILD INSTALLER
REM =========================================================

echo.
echo Step 9: Checking Inno Setup...

where iscc >nul 2>nul

if errorlevel 1 (
    echo [INFO] Inno Setup Compiler not found.
    echo [INFO] Installer build skipped.
    set "step9Status=SKIPPED"
    goto FinalReport
)

echo [OK] Inno Setup found.
echo Compiling installer...

iscc "%rootDir%\installer.iss"

if not "%errorlevel%"=="0" (
    echo [ERROR] Installer build failed.
    set "step9Status=FAILED"
) else (
    echo [OK] Installer compiled successfully.
    set "step9Status=PASSED"
)

REM =========================================================
REM FINAL REPORT
REM =========================================================

:FinalReport

echo.
echo =========================================================
echo                BUILD REPORT: MY-CARTON
echo =========================================================
echo Step 1: Folders Setup      - %step1Status%
echo Step 2: Solution Check     - %step2Status%
echo Step 3: NuGet Restore      - %step3Status%
echo Step 4: MSBuild Check      - %step4Status%
echo Step 5: Project Build      - %step5Status%
echo Step 8: Archive Build      - %step8Status%
echo Step 9: Installer Build    - %step9Status%
echo =========================================================
echo Project Owner: Dmytro Kliuchko
echo Status Date: %date% %time%
echo =========================================================

exit /b 0