1.  Copy Set-WallpaperChanger.ps1 from https://github.com/IronPyrites/WallpaperCollage to C:\Wallpaper.

2.  Copy 1 of the following JSON files to C:\Wallpaper, using the file name applicable to your setting.  Open the file for editing.

		InputsFilePortraitAndLandscape.json
		InputsFilePortraitOnly.json
		InputsFileLandscapeOnly.json
	
	2.1  Update the OutputFolder as needed at the top of the JSON file, or use "C:\\Wallpaper\\Slideshow"
	2.2  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.

	2.3  Update the SourceFolder or copy pictures to the folder applicable to your setting: 
	
		"C:\\Pictures\\PortraitAndLandscape"
		"C:\\Pictures\\PortraitPictures"
		"C:\\Pictures\\LandscapePictures"
	
	2.4  !!! Don't forget to use double "\\" instead of "\" within the folder path in the JSON file.
	2.5  Include the following minimum Portrait and/or Landscape pictures in the folder used above.
	
		PortraitAndLandscape	20 Portrait pictures + 20 Landscape pictures
		PortraitPictures		30 Portrait pictures
		LandscapePictures		30 Landscape pictures
	

3.  Update PowerShell Execution Policy.  This only needs to be done one time.
	3.1  Start -> Windows PowerShell ([right click] -> "Runas Administrator")
	3.2  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
	3.3  Select "A - Yes to all"
	3.4  Close the PowerShell window
	
4.  Open a new PowerShell window (not as an Administrator):  
		Windows [Start] -> PowerShell

5.  Within PowerShell, navigate to C:\Wallpaper, and run the following (if Portrait or Landscape, use the file name from step 3 above):
      
		.\Set-WallpaperChanger.ps1 -InputsJsonFile C:\Wallpaper\InputsFilePortraitAndLandscape.json

    !!! NOTE:  Use the FULL PATH for the JSON file, e.g. C:\Wallpaper\InputsFilePortraitAndLandscape.json, NOT .\InputsFilePortraitAndLandscape.json

6.  Set Windows wallpaper slideshow location
	6.1  Start -> Settings -> Personalization -> Background
	6.2  Disable "Shuffle the picture order"