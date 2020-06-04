@echo off

pushd %~dp0

:cleanOldRst
rem clean old rst
if not exist .\source goto skipCleanOldRst
del /f /s /q .\source\pyoffice.*.rst
del /f /s /q .\source\modules.rst
goto createRst

:skipCleanOldRst
echo "Skip clean old rst"

:createRst
rem create rst
sphinx-apidoc -o .\source ..\src\pyoffice
if %ERRORLEVEL% NEQ 0 goto end

:makeHtml
call make.bat html
if %ERRORLEVEL% NEQ 0 (
    goto cleanBuild
) else (
    goto cleanRstDst
)

:cleanBuild
pushd build
for /d %%f in (*) do (
    echo .\%%f
    rd /s /q .\%%f
)
popd

:cleanRstDst
if not exist .\source goto end
del /f /s /q .\source\pyoffice*.rst
del /f /s /q .\source\modules.rst

:cleanBuildDoctrees
if not exist .\build\doctrees goto end
rd /s /q .\build\doctrees

:end
popd
