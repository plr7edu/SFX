@echo off
setlocal EnableDelayedExpansion

rem === 1. Count subfolders and capture the single folder's name ===
set count=0
for /D %%D in (*) do (
    set /A count+=1
    set "srcFolder=%%D"
)

if NOT "%count%"=="1" (
    echo ERROR: Detected %count% folders. Exactly one is required.
    exit /B 1
)

rem === 2. Prepare output filename ===
set "outExe=%srcFolder%.exe"

rem === 3. Define WinRAR executable path (adjust if needed) ===
if exist "%ProgramFiles%\WinRAR\WinRAR.exe" (
    set "winrar=%ProgramFiles%\WinRAR\WinRAR.exe"
) else (
    rem fallback: assume in PATH
    set "winrar=WinRAR.exe"
)

rem === 4. Create SFX script file ===
echo ;!@Install@!UTF-8!> sfxscript.txt
echo Silent=1>> sfxscript.txt
echo ;!@InstallEnd@!>> sfxscript.txt

rem === 5. Build the SFX archive ===
"%winrar%" a -r -sfx -ibck -inul -o+ -zsfxscript.txt "%outExe%" "%srcFolder%\*"
if ERRORLEVEL 1 (
    echo ERROR: WinRAR reported an error.
    exit /B 2
)

rem === 6. Clean up temporary files ===
del sfxscript.txt

echo SFX archive created successfully: "%outExe%"
endlocal