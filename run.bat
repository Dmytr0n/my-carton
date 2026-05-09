@echo off
REM Встановлюємо кодування для коректного відображення символів
chcp 65001 >nul

set rootDir=%cd%

REM Крок 1: Налаштування структури папок
echo Step 1: Setting up folder structure...
if not exist "%rootDir%\deploy" mkdir "%rootDir%\deploy"
if not exist "%rootDir%\deploy\client" mkdir "%rootDir%\deploy\client"
if not exist "%rootDir%\artefacts" mkdir "%rootDir%\artefacts"

set step1Status=PASSED

REM ШЛЯХИ ДО ФАЙЛІВ (Адаптовано під твій скріншот)
set solutionPath=%rootDir%\src\MyCarton\moy_carton.sln
set clientBuildOutput=%rootDir%\deploy\client
set clientArtifactZipPath=%rootDir%\artefacts\my_carton_build.zip

echo ---------------------------
echo MOY_KARTON BUILD START
echo ---------------------------

REM Крок 2: Перевірка наявності файлу Solution
echo Step 2: Checking solution file...
if not exist "%solutionPath%" (
    echo [ERROR] Solution file not found at: %solutionPath%
    set step2Status=FAILED
    goto FinalReport
) else (
    set step2Status=PASSED
    echo [OK] Solution found.
)

REM Крок 3: Відновлення пакетів NuGet
echo Step 3: Restoring NuGet packages...
REM Переконайтеся, що nuget.exe доступний у PATH або покладіть його поруч із батником
nuget restore "%solutionPath%"
if %errorlevel% neq 0 (
    echo [ERROR] Failed to restore NuGet packages.
    set step3Status=FAILED
    goto FinalReport
) else (
    set step3Status=PASSED
)

REM Крок 4: Перевірка MSBuild
echo Step 4: Checking for MSBuild...
where msbuild >nul 2>nul
if errorlevel 1 (
    echo [ERROR] MSBuild.exe not found. Переконайтеся, що Visual Studio Build Tools встановлені.
    set step4Status=FAILED
    goto FinalReport
) else (
    set step4Status=PASSED
)

REM Крок 5: Збірка проєкту (Release)
echo Step 5: Building project...
msbuild "%solutionPath%" /p:Configuration=Release /p:OutputPath="%clientBuildOutput%" /p:Platform="Any CPU"

if exist "%clientBuildOutput%\moy_carton.exe" (
    echo [OK] Build completed successfully.
    set step5Status=PASSED
) else (
    echo [ERROR] Build failed. EXE file not found in output.
    set step5Status=FAILED
    goto FinalReport
)

REM Крок 8: Архівування результатів
echo Step 8: Creating build artifact archive...
powershell -Command "Compress-Archive -Path '%clientBuildOutput%\*' -DestinationPath '%clientArtifactZipPath%' -Force"
if %errorlevel% neq 0 (
    set step8Status=FAILED
) else (
    set step8Status=PASSED
    echo [OK] Artifact saved at: %clientArtifactZipPath%
)

REM Крок 9: Побудова інсталятора (Inno Setup)
set step9Status=NOT STARTED
where iscc >nul 2>nul
if errorlevel 1 (
    echo [INFO] Inno Setup (iscc) not found. Skipping installer build.
    set step9Status=SKIPPED
) else (
    echo Step 9: Compiling installer...
    iscc "%rootDir%\installer.iss"
    if %errorlevel% neq 0 (
        set step9Status=FAILED
    ) else (
        set step9Status=PASSED
    )
)

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
