@echo off
color 0E
cd %CD%\installers
cls
echo Please wait...
7z x *.exe "AcronisBootableComponentsMediaBuilder.msi" -o*
cls
echo ============
echo Media Builder installers are EXTRACTED!
rem del AcronisBackupAdvanced_*.exe
