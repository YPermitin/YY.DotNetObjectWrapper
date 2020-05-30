@echo off
setlocal

%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe  "..\YY.DotNetObjectWrapper.dll" /codebase
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe  "..\YY.DotNetObjectWrapper.dll" /codebase