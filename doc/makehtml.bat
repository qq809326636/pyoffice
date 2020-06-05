@echo off

pushd %~dp0

:cleanOldRst
rem clean old rst
if not exist %cd%\source goto skipCleanOldRst
del /f /s /q %cd%\source\pyoffice.*.rst
del /f /s /q %cd%\source\modules.rst
rd /s /q %cd%\build\latest

goto createRst

:skipCleanOldRst
echo "Skip clean old rst"

:createRst
rem create rst
sphinx-apidoc -o %cd%\source %cd%\..\src\pyoffice
if %ERRORLEVEL% NEQ 0 goto end

:makeHtml
call make.bat html
if %ERRORLEVEL% NEQ 0 (
    goto cleanBuild
) else (
    ren %cd%\build\html latest
    goto cleanRstDst
)

:cleanBuild
pushd build
for /d %%f in (*) do (
    echo %cd%\%%f
    rd /s /q %cd%\%%f
)
popd

:cleanRstDst
if not exist %cd%\source goto end
del /f /s /q %cd%\source\pyoffice*.rst
del /f /s /q %cd%\source\modules.rst

:cleanBuildDoctrees
if not exist %cd%\build\doctrees goto end
rd /s /q %cd%\build\doctrees
rd /s /q %cd%\build\html

:end
popd
