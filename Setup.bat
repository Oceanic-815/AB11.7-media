@echo off
color 0E
cls
powercfg -X -monitor-timeout-ac 0
powercfg -X monitor-timeout-dc 0
powercfg -X standby-timeout-ac 0
powercfg -X standby-timeout-dc 0
powercfg -X hibernate-timeout-ac 0
powercfg -X hibernate-timeout-dc 0
echo Power management settings were changed
echo Please wait...
"%CD%\bin\python-3.6.0.exe" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 TargetDir="C:\Program Files (x86)\python360"
"C:\Program Files (x86)\python360\Scripts\pip.exe" install "%CD%\bin\pywinauto-0.6.4"
cls
echo Python is installed. 7zip installation starts. Please wait...
msiexec /i "%CD%\bin\7z1701.msi" /quiet /passive /qn /norestart 
setx PATH "%PATH%;C:\Users\test22\AppData\Local\Programs\Python\Python36-32\;C:\Program Files\7-Zip\;C:\Program Files (x86)\7-Zip"
echo 7zip is installed.
echo ============
echo NOW YOU CAN UNZIP INSTALLERS BY RUNNING UNZIP.BAT
echo OR RUN THE PYTHON SCRIPT (which can also unzip the installers)
pause
exit