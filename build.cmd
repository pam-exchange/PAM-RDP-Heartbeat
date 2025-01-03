set AHK_HOME=C:\opt\AutoHotKey-v1

set DIST=.\dist
if exist %DIST% rmdir /S /Q %DIST%
mkdir %DIST%

rem
rem Build pam-rdp-heartbeat from ahk source
rem
set SRC=.\src
set NAME=pam-rdp-heartbeat
"%AHK_HOME%\Compiler\ahk2exe.exe" /base "%AHK_HOME%\Compiler\Unicode 64-bit.bin" /compress 2 /icon "%SRC%\%NAME%.ico" /in "%SRC%\%NAME%.ahk" /out "%DIST%\%NAME%.exe"

rem
rem Copy documentation and other files 
rem 
rem xcopy /S /I /Y Docs %DIST%\Docs
copy /Y LICENSE %DIST%

rem 
rem Package to a zip file
rem
set NAME=PAM-RDP-Heartbeat
if exist %NAME%.zip rm /Y %NAME%.zip
cd %DIST%
zip -r ..\%NAME%.zip
cd ..