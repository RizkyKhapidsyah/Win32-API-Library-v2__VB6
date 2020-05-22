@echo off
IF "%1" == "" goto USAGE
IF "%2" == "" goto USAGE

mc %1.mc
IF errorlevel 1 goto ERRLEVEL
rc -r -fo %1.res %1.rc
IF errorlevel 1 goto ERRLEVEL
link -dll -noentry -out:%2.dll %1.res
IF errorlevel 1 goto ERRLEVEL
del %SystemRoot%\system32\%2.dll
IF errorlevel 1 goto ERRLEVEL
copy %2.dll %SystemRoot%\system32\
IF errorlevel 1 goto ERRLEVEL

:ERRLEVEL
echo.
echo ***
echo *** BUILD FAILED -- ErrorLevel is non-zero!
echo ***
goto FINISH

:USAGE
echo The usage for this file is: build [MESSAGE_FILE] [MESSAGE_DLL]
:FINISH