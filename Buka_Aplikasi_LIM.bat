@echo off
setlocal enabledelayedexpansion

set "HTMLFILE=%USERPROFILE%\Downloads\ketua_lim_app.html"

if not exist "%HTMLFILE%" (
    echo File tidak ditemukan: %HTMLFILE%
    echo Pastikan ketua_lim_app.html ada di folder Downloads.
    pause
    exit /b
)

set "FILEURL=file:///%HTMLFILE:\=/%"
set "FILEURL=!FILEURL: =%%20!"

set "CHROME1=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
set "CHROME2=%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"
set "CHROME3=%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"

if exist "%CHROME1%" (
    start "" "%CHROME1%" --app="%FILEURL%" --window-size=430,900
    exit /b
)
if exist "%CHROME2%" (
    start "" "%CHROME2%" --app="%FILEURL%" --window-size=430,900
    exit /b
)
if exist "%CHROME3%" (
    start "" "%CHROME3%" --app="%FILEURL%" --window-size=430,900
    exit /b
)

set "EDGE1=%PROGRAMFILES(X86)%\Microsoft\Edge\Application\msedge.exe"
set "EDGE2=%PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe"
set "EDGE3=%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"

if exist "%EDGE1%" (
    start "" "%EDGE1%" --app="%FILEURL%" --window-size=430,900
    exit /b
)
if exist "%EDGE2%" (
    start "" "%EDGE2%" --app="%FILEURL%" --window-size=430,900
    exit /b
)
if exist "%EDGE3%" (
    start "" "%EDGE3%" --app="%FILEURL%" --window-size=430,900
    exit /b
)

echo Chrome dan Edge tidak ditemukan. Membuka di browser default...
start "" "%FILEURL%"
