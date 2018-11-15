@echo OFF

where python.exe 1>nul 2>nul
if not errorlevel 1 (
    echo Python executable found on path
    echo.
    python.exe "DATA\InstallPythonInterface.py"
    
)

call :RunScript 2.7
if not errorlevel 1 exit /b

call :RunScript 3.0
if not errorlevel 1 exit /b

call :RunScript 3.1
if not errorlevel 1 exit /b

call :RunScript 3.2
if not errorlevel 1 exit /b

call :RunScript 3.3
if not errorlevel 1 exit /b

call :RunScript 3.4
if not errorlevel 1 exit /b

call :RunScript 2.7 Wow6432Node\
if not errorlevel 1 exit /b

call :RunScript 3.0 Wow6432Node\
if not errorlevel 1 exit /b

call :RunScript 3.1 Wow6432Node\
if not errorlevel 1 exit /b

call :RunScript 3.2 Wow6432Node\
if not errorlevel 1 exit /b

call :RunScript 3.3 Wow6432Node\
if not errorlevel 1 exit /b

call :RunScript 3.4 Wow6432Node\
if not errorlevel 1 exit /b

exit /b

:RunScript
for /f "usebackq skip=2 tokens=1,2*" %%a in (`reg query HKLM\Software\%~2Python\PythonCore\%~1\InstallPath /ve 2^>nul`) do (
    if not "%%c" == "" (
        echo Python %~1 installation: found at %%c
        echo.
        "%%cpython" DATA\InstallPythonInterface.py
        exit /b 0
    ) else (
        echo Python %~1 installation: not found
        exit /b 1
    )
)
pause