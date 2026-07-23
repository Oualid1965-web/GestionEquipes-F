@echo off
setlocal enabledelayedexpansion
title Gestion des Equipes - Construction automatique
color 0A

echo =========================================================
echo   Gestion des Equipes - Construction automatique du .exe
echo =========================================================
echo.
echo Ce script fait TOUT automatiquement, en un seul clic :
echo   - installe Python si necessaire (invisible, sans assistant)
echo   - installe les outils necessaires
echo   - construit GestionEquipes.exe
echo.
echo Aucune action de votre part n'est necessaire. Patientez...
echo.

set "PYEXE="

where python >nul 2>nul
if %errorlevel%==0 (
    for /f "delims=" %%P in ('where python') do (
        if not defined PYEXE set "PYEXE=%%P"
    )
)

if not defined PYEXE (
    echo [1/4] Python n'est pas present sur cet ordinateur.
    echo       Telechargement et installation automatique en cours...
    echo       ^(cela prend 1 a 3 minutes, aucune fenetre ne s'affichera^)
    echo.

    set "PYVER=3.12.13"
    set "PYURL=https://www.python.org/ftp/python/!PYVER!/python-!PYVER!-amd64.exe"
    set "PYINSTALLER_EXE=%TEMP%\python-installer.exe"

    powershell -NoProfile -Command "try { Invoke-WebRequest -Uri '!PYURL!' -OutFile '!PYINSTALLER_EXE!' -UseBasicParsing } catch { exit 1 }"
    if not exist "!PYINSTALLER_EXE!" (
        echo.
        echo [ERREUR] Le telechargement de Python a echoue.
        echo Verifiez votre connexion internet, ou installez Python manuellement
        echo depuis https://www.python.org/downloads/ ^(cochez "Add python.exe to PATH"^)
        echo puis relancez ce script.
        echo.
        pause
        exit /b 1
    )

    "!PYINSTALLER_EXE!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=0
    echo       Installation de Python terminee.
    echo.

    for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
        if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
    )
    if not defined PYEXE (
        for /d %%D in ("%ProgramFiles%\Python3*") do (
            if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
        )
    )
    if not defined PYEXE (
        echo [ERREUR] L'installation de Python n'a pas ete detectee correctement.
        echo Installez Python manuellement depuis https://www.python.org/downloads/
        echo ^(cochez bien "Add python.exe to PATH"^), puis relancez ce script.
        pause
        exit /b 1
    )
) else (
    echo [1/4] Python detecte : !PYEXE!
)

echo [2/4] Mise a jour de pip...
"!PYEXE!" -m pip install --upgrade pip >nul 2>nul

echo [3/4] Installation des dependances ^(pyinstaller, openpyxl^)...
"!PYEXE!" -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERREUR] L'installation des dependances a echoue. Verifiez votre connexion internet.
    pause
    exit /b 1
)

echo [4/4] Construction de l'executable Windows ^(.exe^)...
"!PYEXE!" -m PyInstaller --onefile --windowed --name "GestionEquipes" app.py
if %errorlevel% neq 0 (
    echo [ERREUR] La construction a echoue. Voir le message ci-dessus.
    pause
    exit /b 1
)

copy /Y "dist\GestionEquipes.exe" "GestionEquipes.exe" >nul

echo.
echo =========================================================
echo   TERMINE !
echo   Votre application se trouve ici :
echo   %cd%\GestionEquipes.exe
echo.
echo   Vous pouvez maintenant deplacer ce fichier .exe
echo   n'importe ou sur l'ordinateur ^(bureau, dossier partage, etc.^)
echo   et creer un raccourci. Aucune connexion internet, aucun Python
echo   et aucune installation ne sont necessaires pour l'utiliser.
echo =========================================================
echo.
pause
