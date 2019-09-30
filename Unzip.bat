@echo off
color 0E
cd %CD%\installers
cls
echo Extracting using Unzip.bat. Please wait...
7z x *.exe "AcronisBootableComponentsMediaBuilder.msi" -o*
cls
echo ============
echo Unzip.bat execution is finished!
del AcronisBackupAdvanced_*.exe
