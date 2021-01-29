Write-Output 'Custom PowerShell profile in effect!'
PowerShell.exe -ExecutionPolicy Bypass -Command "& '%~dpn0.ps1'"
@ECHO OFF
