@ECHO OFF
:BEGIN

:COPY
xcopy \\snap\share1\ASN\Script\*.* .\ /D /E /C /Q /H /R /Y /K 

:RUN
cmd /c powershell.exe -file .\AtechASN.ps1 > script.log