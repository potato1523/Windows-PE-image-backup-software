@echo off
echo ============================================
echo  WinPE Image Tool - Build (MSVC)
echo ============================================
echo.

REM Try to find Visual Studio
where cl >nul 2>&1
if %ERRORLEVEL% EQU 0 goto :BUILD

REM Try VS2022 Developer Command Prompt
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" (
    call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
    goto :BUILD
)
if exist "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" (
    call "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
    goto :BUILD
)
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars64.bat" (
    call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars64.bat"
    goto :BUILD
)

echo [ERROR] Visual Studio or Build Tools not found.
echo Please run from "Developer Command Prompt" or install Build Tools.
pause
exit /b 1

:BUILD
echo Compiling...
cl /O2 /W4 /EHsc /DUNICODE /D_UNICODE /std:c++17 ^
    main.cpp ^
    /Fe:WinPE-ImageTool.exe ^
    /link /SUBSYSTEM:WINDOWS ^
    user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Build successful: WinPE-ImageTool.exe
    echo.
    REM Clean up intermediate files
    del /q main.obj 2>nul
) else (
    echo.
    echo [ERROR] Build failed.
)

pause
