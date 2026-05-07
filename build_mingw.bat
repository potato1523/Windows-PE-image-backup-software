@echo off
echo ============================================
echo  WinPE Image Tool - Build (MinGW-w64)
echo ============================================
echo.

where g++ >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] g++ not found. Please install MinGW-w64 and add to PATH.
    pause
    exit /b 1
)

echo Compiling...
g++ -O2 -std=c++17 -DUNICODE -D_UNICODE -mwindows ^
    main.cpp ^
    -o WinPE-ImageTool.exe ^
    -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshell32 -lole32 ^
    -static -static-libgcc -static-libstdc++

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Build successful: WinPE-ImageTool.exe
) else (
    echo.
    echo [ERROR] Build failed.
)

pause
