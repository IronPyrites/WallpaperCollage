# Step 1:  Update PowerShell Execution Policy
1. Start -> Windows PowerShell ([right click] -> "Runas Administrator")
2. Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
3. Select "A - Yes to all"


# Step 2:  Stage Files
Copy the following files into a single folder, e.g. C:\Wallpaper
```
Set-WallpaperChanger.ps1
WallpaperChangerTask_PowerShell_5.cmd
WallpaperChangerTask_PowerShell_7.cmd
WallpaperChanger.xml
InputsFile.json
```


# Step 3:  Configuration
1. **Read usage notes at the top of Set-WallpaperChanger.ps1.**

    Open InputsFile.json with a proper JSON file editor, update global parameters (top), SourceFolderGroups, SourceFolderGroupSets, and WallpaperLayoutWeightedList, using the examples provided. 

    VSCode is free and works well for editing JSON files:  https://code.visualstudio.com/docs/?dv=win64user

2. Update parameters as needed in WallpaperChangerTask.cmd.

**Note:  Parameters may be changed in InputsFile.json at any time, even while the Task is running in Windows Task scheduler.  Parameters will be implemented immediately after saving InputsFile.json, without the need to restart the script (or Task in Windows Task Scheduler).  However, there will be a delay due to TimeBetweenCreation (max wait time is [NumberThreads] x [TimeBetweenCreation]).**


# Step 4:  Running Set-WallpaperChanger.ps1
Notes about PowerShell version:
- PowerShell version 7 is required for multithreading, and can be installed from: https://github.com/PowerShell/PowerShell/releases/download/v7.5.1/PowerShell-7.5.1-win-x64.msi. 
- Multithreading is useful if Wallpaper image creation times are longer than the slideshow interval.

1. A single parameter is used, since the remainder of input parameters are stored in InputsFile.json.

2. **!!! IMPORTANT !!!** Use the FULL path of InputsFile.json, e.g. C:\Wallpaper\InputsFile.json, since the full path will be self-referenced during runtime.
```
.\Set-WallpaperChanger.ps1 -InputsJsonFile C:\Wallpaper\InputsFile.json
```

# Step 5:  Scheduled Task (Optional, Recommended)
1. Start -> Run "Task Scheduler." Import WallpaperChanger.xml into Windows Task Scheduler and modify as needed.
2. Use WallpaperChangerTask_PowerShell_5.cmd for PowerShell version 5 (Windows default), or WallpaperChangerTask_PowerShell_7.cmd for PowerShell version 7.  Update the files location (Step 2) in either CMD file used.
3. Set the Task Scheduler task to run under a different Administrator account (e.g. local Administrator) to eliminate popup windows.
  
4. Reboot your Computer.


# Step 6:  Windows - Set Wallpaper Location
1. Start -> Settings -> Personalization -> Background
2. Disable "Shuffle the picture order"


# Step 7:  Windows - Custom Slideshow Interval
Windows Settings -> Personalization -> Background has a limited set of values.  
The following may be run in PowerShell to upate the Windows registry for more flexibility.  A value of 10000 (microseconds) is 10 seconds, and is the minimum for Windows 11.
Requires a logoff/logon or a reboot for the setting to take effect.

Run in PowerShell:
```
Get-Item 'HKCU:Control Panel\Personalization\Desktop Slideshow' | Set-ItemProperty -Name Interval -Value 10000
```