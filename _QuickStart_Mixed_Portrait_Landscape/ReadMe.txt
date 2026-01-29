1.  Copy this simplified InputsFile.json file to C:\Wallpaper.

2.  Copy Set-WallpaperChanger.ps1 from https://github.com/IronPyrites/WallpaperCollage to your local C:\Wallpaper folder.

3.  Open  InputsFile.json for editing:
      3.1  Update the OutputFolder as needed at the top of InputsFile.json, or use "C:\\Wallpaper\\Slideshow"
      3.2  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.

      3.3  Update the SourceFolder or copy pictures to "C:\\Pictures\\MixedPortraitLandscape"
      3.4  Make sure you include at least 30 Portrait and 30 Landscape pictures in SourceFolder.
      3.5  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.

4.  Open a PowerShell window:  
      Windows [Start] -> PowerShell

5.  Update PowerShell Execution Policy
      5.1  Start -> Windows PowerShell ([right click] -> "Runas Administrator")
      5.2  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
      5.3  Select "A - Yes to all"

6.  Within PowerShell, navigate to C:\Wallpaper, and run the following:
      
      .\Set-WallpaperChanger.ps1 -InputsJsonFile C:\Wallpaper\InputsFile.json

    !!! NOTE:  Use the FULL PATH for InputsFile.json, e.g. C:\Wallpaper\InputsFile.json, NOT .\InputsFile.json

7.  See _QuickStartGuide.md steps 6 and 7 for setting Windows to use OutputFolder (above) as the slideshow location.