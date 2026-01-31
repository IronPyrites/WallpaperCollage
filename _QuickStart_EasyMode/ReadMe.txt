1.  Copy this simplified InputsFileLandscape.json file to C:\Wallpaper.

2.  Copy Set-WallpaperChanger.ps1 from https://github.com/IronPyrites/WallpaperCollage to your local C:\Wallpaper folder.

3.  Open 1 of the following JSON files for editing, using the file name applicable to your setting:

		InputsFilePortraitAndLandscape.json
		InputsFilePortraitOnly.json
		InputsFileLandscapeOnly.json
	
	3.1  Update the OutputFolder as needed at the top of the JSON file, or use "C:\\Wallpaper\\Slideshow"
	3.2  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.

	3.3  Update the SourceFolder or copy pictures to the folder applicable to your setting: 
	
		"C:\\Pictures\\PortraitAndLandscape"
		"C:\\Pictures\\PortraitPictures"
		"C:\\Pictures\\LandscapePictures"
	
	3.4  Make sure you include at least 30 Landscape pictures in SourceFolder.
	3.5  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.

4.  Update PowerShell Execution Policy.  This only needs to be done one time.
	4.1  Start -> Windows PowerShell ([right click] -> "Runas Administrator")
	4.2  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
	4.3  Select "A - Yes to all"
	4.4  Close the PowerShell window
	
5.  Open a new PowerShell window (not as an Administrator):  
		Windows [Start] -> PowerShell

6.  Within PowerShell, navigate to C:\Wallpaper, and run the following (if Portrait or Landscape, use the file name from step 3 above):
      
		.\Set-WallpaperChanger.ps1 -InputsJsonFile C:\Wallpaper\InputsFilePortraitAndLandscape.json

    !!! NOTE:  Use the FULL PATH for the JSON file, e.g. C:\Wallpaper\InputsFilePortraitAndLandscape.json, NOT .\InputsFilePortraitAndLandscape.json

7.  Set Windows wallpaper slideshow location
	7.1  Start -> Settings -> Personalization -> Background
	7.2  Disable "Shuffle the picture order"