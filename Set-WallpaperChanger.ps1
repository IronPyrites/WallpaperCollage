<#
.SYNOPSIS
Creates Wallpaper collage files using combinations of Landscape and Portrait images.

All parameters are stored in InputsFile.json.  Updates to InputsFile.json are implemented immediately, without the need to restart the script.

Author:   Dustin Fraser
Version:  20260125


.DESCRIPTION

Script Parameters:

	InputsJsonFile - Script only parameter and the file containing all script parameters, the script will not run without it.  All parameters are interpreted by the Get-InputJsonFileData function.
		- It is STRONGLY recommended that VSCode or other JSON file editor be used to edit the JSON.  https://code.visualstudio.com/docs/?dv=win64user
		- A malformed JSON file will result in the script not running, since all parameters are derived form InputsJsonFile.json.
		- VSCode or other JSON editors will warn about formatting mistakes such as a single "\" instead of "\\".  Again, please use a JSON editor.

	Thread
		- Do not use, used internally in support of NumberThreads, and will result in a single Wallpaper file being updated
		
	AutoMultiThreadActive
		- Do not use, used internally in support of AutoMultiThread.

InputsFile.json Parameters:

	OutputFolder - Windows Wallpaper location.  
		- All folder backslash "\" characters need to be "escaped" in the JSON file, e.g. C:\Wallpaper\OutputFolder" is stored as "C:\\Wallpaper\\OutputFolder". i.e. replace "\" with "\\"
		- Update Windows settings for the OutputFolder location:

			[Start] -> Run -> Settings -> Personalization -> Background
				Disable "Shuffle the picture order"
				
	TimeBetweenCreation
		- Should be the same (or 1-2 less) as the Windows slideshow interval (see below).
		- Note that large image files and complex layouts may take 60+ seconds to create.  If so, and duplicate slideshow Windows backgrounds are a regular occurrence, consider using NumberThreads (below). 

    FolderEnumerationTimeLimit
        - SourceFolder enumeration will be limited to FolderEnumerationTimeLimit seconds to prevent runtime from getting stuck in the enumeration phase; remaining SourceFolders will be truncated.

	NumberThreads
        - Requires PowerShell 7 or higher:  https://github.com/PowerShell/PowerShell/releases/download/v7.5.1/PowerShell-7.5.1-win-x64.msi
		- Number of parallel threads, each of which will update its own Wallpaper image file.
		- Parameter values of 1 through 6 are supported, invalid values will be replaced with "2"
		- Recommended setting, using 4-6 threads.
		
	AutoMultiThread
        - Requires PowerShell 7 or higher:  https://github.com/PowerShell/PowerShell/releases/download/v7.5.1/PowerShell-7.5.1-win-x64.msi
		- Values of "true" or "false" (without quotes)
		- If NumberThreads is set to 1 and AutoMultiThread is "true," threads will automatically be increased and decreased if creation time is exceeding TimeBetweenCreation
		- Not recommended, consider using NumberThreads for more consistent behavior.

	IncludeRecurseSubdirectoriesTagFile 
		- Only include folders (and their subdirectories) which include the IncludeRecurseSubdirectoriesTagFile file (e.g. WP.include file (empty)).
		- Not required, if a IncludeRecurseSubdirectoriesTagFile file does not exist, the entire SourceFolder and subdirectories will be included by default.
		- However, once a IncludeRecurseSubdirectoriesTagFile is included in one or more folders, all other parallel folders/subdirectory chains will be excluded.

	ExcludeRecurseSubdirectoriesTagFile 
		- Exclude folders (and their subdirectories) which include the ExcludeRecurseSubdirectoriesTagFile file (e.g. WP.exclude file (empty)).
		- May also be used to prune (exclude) subdirectories of IncludeRecurseSubdirectoriesTagFile parent folders.	

	PauseOnProcess 
		- List of Windows processes which will pause Wallpaper image creation, as defined by "Get-Process | Select-Object ProcessName -Unique" (PowerShell).
		- E.g. ["powerpnt","steam"] to conserve CPU while presenting or playing games.

    
    SourceFolderGroups (InputsFile.json property):
	
		Description - Cosmetic value, may be updated freely.
		
		SourceFolders - May be a single folder or a list.
			- All folder backslash "\" characters need to be "escaped" in the JSON file, e.g. C:\Wallpaper\Sourcefolder is stored as "C:\\Wallpaper\\Sourcefolder". i.e. replace "\" with "\\"

		SourceFolderGroup 
			- Integer identifier
			- Should be unique, but will be deduplicated in either case.
			- Weight, ZoomPercent, and RandomFlipHorizontal override settings here will only be used here if the WallpaperLayoutWeightedList GroupName is set to:
			    SplitGroupOnSourceFolder = true (see below under section WallpaperLayoutWeightedList).
			
            Weight (SourceFolder override)
                - Integer multiplier (1,2,3...) indicating how often a Folder should be represented.  
			    - For example, image files in Folders with a Weight of 10 will be displayed 10x as often as image files in a Folder with a Weight of 1.
			
            ZoomPercent (SourceFolder override)
                - See WallpaperLayoutWeightedList notes below (may be set more globally for a given GroupName)

            RandomFlipHorizontal (SourceFolder override)
                - See WallpaperLayoutWeightedList notes below (may be set more globally for a given GroupName)

            SplitOnLeafFolder
                - CAUTION when using this setting, a [Large numbers of LeafFolders] x [SourceFolder Weight values] will take a long time to enumerate
                  Recommendation:  If enabled, use a SourceFolder Weight of 1.
                  Enumeration will be limited to FolderEnumerationTimeLimit seconds (see above).
                - Split a parent SoureFolder into multiple picture-containing subdirectory SourceFolders with the longest paths, i.e. "LeafFolders".
                - This setting is intended to create WallpaperImage files based on a single folder, rather than on an entire directory and subdirectory tree.
                - Picture files contained within branch directories (upstream of longest-path picture-containing subdirectories) will not be included.
                - NOTE that ExcludeRecurseSubdirectoriesTagFile (see above), meant to prune subdirectories (branches) of parent SourceFolders,
                  will only be applied if set explicitly on LeafFolders. IncludeRecurseSubdirectoriesTagFile is enforced only if one (1) or more LeafFolders
                  contain IncludeRecurseSubdirectoriesTagFile; otherwise, all LeafFolders will be included.
                - Values of "true" or "false" (without quotes).

	
    SourceFolderGroupSets (InputsFile.json property):

		SourceFolderGroups
			- Group of 1 or more SourceFolderGroup integer identifiers (from above)
			
		SourceFolderGroupSet
			- MUST be a unique integer identifier
			- Used to associate SourceFolderGroups with WallpaperLayoutWeightedList GroupNames (below)

	
    WallpaperLayoutWeightedList (InputsFile.json property):
	
		GroupName - Cosmetic name description for WallpaperLayout setting combinations, with the exception of the "Default" GroupName which should be preserved.
		
		WallpaperLayout 
			- Image file collage layout within the WallpaperImage file.  
			- Valid layout values are provided in $WallpaperLayoutCountHash below, e.g. "Landscape_1_Base6Thirds"
			- Its best to just experiment rather than attempting to describe them here.
		
		Weight - Similar to Weight used on SourceFolders (above), weighted representation of a given GroupName.  Must be a positive integer.
		
		SourceFolderGroupSets 
			- List of SourceFolderGroupSet(s) from above.
			- Integer list, e.g. [2], [1,2,4]
		
		SplitGroupOnSourceFolderGroupSets 
			- If SplitGroupOnSourceFolderGroupSets = True (recommended), auto-create new GroupNames so that each Wallpaper image does not contain a mix of SourceFolderGroups.
			- Not applicable to the "Default" GroupName.
			
		SplitGroupOnSourceFolder
			- If SplitGroupOnSourceFolder = True (recommended)
				- Auto-create new GroupNames so that each Wallpaper image does not contain a mix of Folders within a SourceFolders list.
				- Enables per Folder overrides of Weight, ZoomPercent, and RandomFlipHorizontal.

		IncludeSubdirectories - Also include image files from all subdirectories (subfolders) of a given SourceFolder.

		ZoomPercent 
			- Zooms in on the center of images, effectively cropping them.  Useful for photographers who chronically over-panoramaize photos, leaving the subject too small.
			- Must be an integer value greater than or equal to 100.
			- A ZoomPercent of 100 will not zoom in on or crop the image (recommended).

        RandomFlipHorizontal - Randomly flip SourceFolder image files horizontally.  Provides variety for picture art, but text, scenery, etc. will be confusingly mirrored.

        AlphaNumeric
			- Add pictures to Wallpaper images from SourceFolders in AlphaNumeric order.  
			- Good for continuity, such as displaying vacation pictures in order (assuming the dates follow alphanumerical order).

		RandomSequence - A starting image is randomly selected selected within the SourceFile list, with subsequent images followed alphanumerically for the duration of the Wallpaper image creation process.

        Portrait_Landscape_SurplusBurnoff
            - Image files are scanned for Portrait or Landscape designation ONCE per SourceFolder enumeration to conserve disk and CPU resources.
            - If Portrait file is requested but a Landscape file is identified, the pre-designated Landscape file name is cached for later use (and vice versa).
            - In the event there is a large surplus of Portrait and/or Landscape cached image file names, surpluse files will be burned off using Portrait or Landscape heavy layouts.
            - Recommend setting this to True.

        SurplusCandidate_Portrait - Layout for burning off surplus Portrait images.  Recommend using "Portrait_3_Elements_Base3_Even"

        SurplusCandidate_Landscape - Layout for burning off surplus Landscape images.  Recommend using "Landscape_2_Elements_Thirds"


Windows Settings:

	Wallpaper Refresh Time:

		Personalization -> Background has a limited set of values.  The following example and registry key may be run in PowerShell to upate the Windows registry for more flexibility.  
		10 seconds is "-Value 10000" (microseconds), and is the minimum for Windows 11.

		[Start] -> Run  -> Powershell:
		
			Get-Item 'HKCU:Control Panel\Personalization\Desktop Slideshow' | Set-ItemProperty -Name Interval -Value 10000

		Updating this setting normally requires a logoff/logon or a reboot for the setting to take effect, but should also refresh by visiting:
		
			[Start] -> Settings -> Personalization -> Background


Recommended Settings:
   
    Set Wallpaper JPEG transcoding to 100% quality, C:\Users\<UserName>\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper.  This requires a reboot to take effect.
    
		# PowerShell
		if ((Get-Item 'HKCU:Control Panel\Desktop\').GetValueNames() -notcontains 'JPEGImportQuality') {Get-Item 'HKCU:Control Panel\Desktop\' | New-ItemProperty -Name JPEGImportQuality -Type DWord -Value 100}
		else {Get-Item 'HKCU:Control Panel\Desktop\' | Set-ItemProperty -Name JPEGImportQuality -Value 100}

    Disable Desktop Wallpaper Shuffling (Wallpaper images will be used in numerical order)
	
		# PowerShell
		Get-Item 'HKCU:Control Panel\Personalization\Desktop Slideshow\' | Set-ItemProperty -Name Shuffle -Value 00000000
		
		# Windows UI
		[Windows Start] -> System -> Background -> Choose your desktop background -> Shuffle the picture order (set to "Off")
	

Troubleshooting:  

	1.  Random black desktops, try the following:

		If Desktop Windows Manager (dwm.exe) is crashing, extend GPU Timeout Detection from the default of 2 seconds
		
			# PowerShell
			if ((Get-Item 'HKLM:System\CurrentControlSet\Control\GraphicsDrivers\').GetValueNames() -notcontains 'TdrDelay') {Get-Item 'HKLM:System\CurrentControlSet\Control\GraphicsDrivers\' | New-ItemProperty -Name TdrDelay -Type DWord -Value 5}
			else {Get-Item 'HKLM:System\CurrentControlSet\Control\GraphicsDrivers\' | Set-ItemProperty -Name TdrDelay -Value 5}

		Turn Off "Hardware-accelerated GPU scheduling" via the Windows UI:
		[Windows Start] -> System -> Display Settings -> Graphics -> Change default graphics settings -> Hardware-accelerated GPU scheduling
		
	2.  Wallpaper images stopped being created after InputsFile.json was updated:
	
		Check InputsFile.json for errors using a JSON file editor.
		
	3.  Windows slideshow background images are occastionally repeating:
	
		Try using NumberThreads with a value between 4 and 6 (max).

    4.  Unhandled exception case:  A folder with all Portrait images is erroneously added to a layout with Landscape images (and vice versa).
        
        If NumberThreads is also used, creation threads may never stabilize long enough to use Portrait_Landscape_SurplusBurnoff.
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$InputsJsonFile,
    [Parameter(Mandatory=$false)]
    [Int32]$Thread,
    [Parameter(Mandatory=$false)]
    [string]$AutoMultiThreadActive
)


############################### Static Variables ###############################

# All Wallpaper layouts are intended for a 16:9 aspect ratio, including current 4K (3840 x 2160) and future 8K (7680 x 4320)
$WallpaperDimensions = @{}
$WallpaperDimensions.'WallpaperWidth' = 3840
$WallpaperDimensions.'WallpaperHeight' = 2160

$WallpaperLayoutCountHash = @{}
$WallpaperLayoutCountHash.'Landscape_1_Base6Thirds' = 4
$WallpaperLayoutCountHash.'Landscape_2_Elements_Thirds' = 4
$WallpaperLayoutCountHash.'Landscape_3_Elements_Base3' = 6
$WallpaperLayoutCountHash.'Portrait_1_2x2_3x3' = 6
$WallpaperLayoutCountHash.'Portrait_2_NarrowThirds' = 4
$WallpaperLayoutCountHash.'Portrait_3_Elements_Base3_Narrow' = 6
$WallpaperLayoutCountHash.'Portrait_4_Elements_Base3_Even' = 6
$WallpaperLayoutCountHash.'Mixed_1_6xElements_Base6Thirds_15Px6L' = 4
$WallpaperLayoutCountHash.'Mixed_2_Thirds' = 4
$WallpaperLayoutCountHash.'Mixed_3_Element_Base3' = 6
$WallpaperLayoutCountHash.'Mixed_4_LargeLandscape' = 4
$WallpaperLayoutCountHash.'Magazine' = 4

###############################################################################


function Get-InputJsonFileData {

    Param(
        [Parameter(Mandatory=$false)]
        [string]$InputsJsonFile,
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [hashtable]$WallpaperDimensions,
        [Parameter(Mandatory=$false)]
        [hashtable]$WallpaperLayoutCountHash,
        [Parameter(Mandatory=$false)]
        [Int32]$Thread,
        [Parameter(Mandatory=$false)]
        [string]$AutoMultiThreadActive
    )

    # Save $DataHashtable.InputsJsonFile and $DataHashtable.WallpaperLayoutCountHash before re-initializing the hashtable, only needed for bootstrap in the main body
    if ($DataHashtable.InputsJsonFile) {$InputsJsonFile = $DataHashtable.InputsJsonFile.Clone()}
    if ($DataHashtable.WallpaperDimensions) {$WallpaperDimensions = $DataHashtable.WallpaperDimensions.Clone()}
    if ($DataHashtable.WallpaperLayoutCountHash) {$WallpaperLayoutCountHash = $DataHashtable.WallpaperLayoutCountHash.Clone()}
    if ($DataHashtable.AutoMultiThreadActive) {$AutoMultiThreadActive = [System.Convert]::ToBoolean(([string]$DataHashtable.AutoMultiThreadActive).Clone())}
    
    # PowerShell cannot clone Integer values, copying will lead to access method issues when $DataHashtable is returned from function, we will convert back to [Int32] later
    if ($DataHashtable.PictureNameOverride) {$PictureNameOverride = ([string]$DataHashtable.PictureNameOverride).Clone()}
    if ($DataHashtable.Thread) {$Thread = ([string]$DataHashtable.Thread).Clone()}

    $DataHashtable = @{}

    # Add or re-add back to DataHashtable after initialization
    $DataHashtable.WallpaperDimensions = $WallpaperDimensions
    $DataHashtable.WallpaperLayoutCountHash = $WallpaperLayoutCountHash
    if ($PictureNameOverride) {$DataHashtable.Add('PictureNameOverride', [Int32]$PictureNameOverride)} else {$DataHashtable.Add('PictureNameOverride', 0)}
    if ($AutoMultiThreadActive) {$DataHashtable.Add('AutoMultiThreadActive', $true)} else {$DataHashtable.Add('AutoMultiThreadActive', $false)}
    if ($Thread) {$DataHashtable.Add('Thread', [Int32]$Thread)}
    
    # If the full file path of -InputsJsonFile was not provided, try appending the path the script is being run from
    if ([System.IO.File]::Exists($InputsJsonFile) -and ($InputsJsonFile.Split('\')[0] -ne '.'))
    {
        $DataHashtable.InputsJsonFile = $InputsJsonFile
    }
    elseif ($InputsJsonFile.Split('\')[0] -eq '.')
    {
        $DataHashtable.InputsJsonFile = "${PSScriptRoot}\$($InputsJsonFile.Split('\')[1])"
    }
    elseif ([System.IO.File]::Exists("${PSScriptRoot}\${InputsJsonFile}"))
    {
        $DataHashtable.InputsJsonFile = "${PSScriptRoot}\${InputsJsonFile}"
    }
    else
    {
        $DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path ".\ErrorLog.txt" -Value "${DateTime} The full file path for -InputsJsonFile $InputsJsonFile is invalid or cannot be found, exiting..."

        exit
    }

    if ($PSVersionTable.PSVersion.Major -ge 7)
    {
        $InputsJsonObject = Get-Content -Path $DataHashtable.InputsJsonFile -Raw | ConvertFrom-Json
    }
    else {$InputsJsonObject = Invoke-RestMethod -Method “GET” -uri $DataHashtable.InputsJsonFile}

    $DataHashtable.OutputFolder = [string]$InputsJsonObject.OutputFolder
    if ($DataHashtable.OutputFolder -and (-not (Test-Path -Path $DataHashtable.OutputFolder)))
    {
        $void = New-Item -ItemType Directory -Path $DataHashtable.OutputFolder
    }
    
    [Int32]$DataHashtable.TimeBetweenCreation = $InputsJsonObject.TimeBetweenCreation
    [Int32]$DataHashtable.FolderEnumerationTimeLimit = $InputsJsonObject.FolderEnumerationTimeLimit
    [array]$DataHashtable.PauseOnProcess = $InputsJsonObject.PauseOnProcess
    [string]$DataHashtable.IncludeRecurseSubdirectoriesTagFile = $InputsJsonObject.IncludeRecurseSubdirectoriesTagFile
    [string]$DataHashtable.ExcludeRecurseSubdirectoriesTagFile = $InputsJsonObject.ExcludeRecurseSubdirectoriesTagFile
    
    # Retain NumberThreads if AutoMultiThreadActive
    if ($NumberThreads -and ($InputsJsonObject.NumberThreads -lt 2))
    {
        $DataHashtable.NumberThreads = [Int32]$NumberThreads
    }
    else
    {
        $DataHashtable.NumberThreads = [Int32](@([Int32]$InputsJsonObject.NumberThreads,6) | measure -Minimum).Minimum  # Maximum of 6 threads
        $DataHashtable.NumberThreads = [Int32](@($DataHashtable.NumberThreads,1) | measure -Maximum).Maximum  # Minimum of 1 thread
    }
    $DataHashtable.AutoMultiThread = [System.Convert]::ToBoolean($InputsJsonObject.AutoMultiThread)
        
    $DataHashtable.WallpaperLayoutWeightedList = @{}
    $SourceFolderGroupsObject = [PSCustomObject]$InputsJsonObject.SourceFolderGroups
    $SourceFolderGroupSetsObject = [PSCustomObject]$InputsJsonObject.SourceFolderGroupSets
    $WallpaperLayoutWeightedListObject = [PSCustomObject]$InputsJsonObject.WallpaperLayoutWeightedList

    # Ingest InputsJsonFile SourceFolderGroups
    $SourceFolderGroupsJsonHashtable = @{}
    $DataHashtable.SplitOnLeafFolder = $false
    
    foreach ($SourceFolderGroupItem in $SourceFolderGroupsObject)
    {
        if ($SourceFolderGroupsJsonHashtable.([Int32]$SourceFolderGroupItem.SourceFolderGroup))
        {
            $DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path ".\ErrorLog.txt" -Value "SourceFolderGroup $($SourceFolderGroupItem.SourceFolderGroup) is non-unique, multiple instances cannnot be added..."
            
            continue
        }
        else
        {
            if ($SourceFolderGroupItem.SourceFolders.Count -le 0) {continue}
            if (($SourceFolderGroupItem.SourceFolders.Weight | measure -Maximum).Maximum -le 0) {continue}
            
            $SourceFolderArray = [System.Collections.Generic.List[[PSCustomObject]]]@()
            
            foreach ($FolderItem in $SourceFolderGroupItem.SourceFolders)
            {
                if ($FolderItem.Folder)
                {
                    # Remove any final backslash "\" characters from folder names
                    if ($FolderItem.Folder.Remove(0, ($FolderItem.Folder.Length - 1)) -eq '\')
                    {
                        $FolderItem.Folder = $FolderItem.Folder.Substring(0, $FolderItem.Folder.Length - 1)
                    }

                    # Split SourceFolders with SplitOnLeafFolder=$true into LeafFolders
                    if ($FolderItem.SplitOnLeafFolder)  
                    {
                        if (Test-Path -Path $FolderItem.Folder)
                        {
                            [array]$LeafFolders = (Get-ChildItem -Path "$($FolderItem.Folder)\*" -Include *.BMP, *.GIF, *.EXIF, *.JPG, *.JPEG, *.PNG, *.TIFF -Recurse).DirectoryName | Select-Object -Unique

                            if ($LeafFolders.Count -gt 0)
                            {
                                foreach ($LeafFolder in $LeafFolders)
                                {
                                    # Do not include branch Folders
                                    if ($LeafFolders -like "$($LeafFolder)\*")
                                    {
                                        continue
                                    }
                            
                                    # PSCustomObject does not support Clone(), serializing and deserializing to break object linkings
                                    $FolderItemSerialized = ConvertTo-Json -InputObject $FolderItem
                                    $LeafFolderItem = ConvertFrom-Json -InputObject $FolderItemSerialized
                            
                                    $LeafFolderItem.Folder = $LeafFolder
                            
                                    $SourceFolderArray.Add($LeafFolderItem)
                                }
                            }
                            else
                            {
                                $SourceFolderArray.Add($FolderItem)

                                continue
                            }

                            if ($DataHashtable.IncludeRecurseSubdirectoriesTagFile)
                            {
                                # If one (1) or more LeafFolders contain IncludeRecurseSubdirectoriesTagFile, only include folders with IncludeRecurseSubdirectoriesTagFile; 
                                # else, include all LeafFolders as the default

                                if (Get-ChildItem -Path $FolderItem.Folder -Include $DataHashtable.IncludeRecurseSubdirectoriesTagFile -Recurse)
                                {
                                    $LeafFolderIncludeFileFolderArray = [System.Collections.Generic.List[[PSCustomObject]]]@()

                                    foreach ($LeafFolder in $SourceFolderArray)
                                    {
                                        if (Get-ChildItem -Path $LeafFolder.Folder -Include $DataHashtable.IncludeRecurseSubdirectoriesTagFile -Recurse)
                                        {
                                            $LeafFolderIncludeFileFolderArray.Add($LeafFolder)
                                        }
                                    }

                                    $SourceFolderArray = $LeafFolderIncludeFileFolderArray
                                }
                            }

                            if ($DataHashtable.ExcludeRecurseSubdirectoriesTagFile)
                            {
                                # Exclude LeafFolders which contain ExcludeRecurseSubdirectoriesTagFile

                                if (Get-ChildItem -Path $FolderItem.Folder -Include $DataHashtable.ExcludeRecurseSubdirectoriesTagFile -Recurse)
                                {
                                    foreach ($LeafFolder in $SourceFolderArray)
                                    {
                                        if (Get-ChildItem -Path $LeafFolder.Folder -Include $DataHashtable.ExcludeRecurseSubdirectoriesTagFile -Recurse)
                                        {
                                            $SourceFolderArray = $SourceFolderArray | Where-Object {$_.Folder -ne $LeafFolder.Folder}
                                        }
                                    }
                                }
                            }
                        }
                        else  # SourceFolder is offline, LeafFolders cannot be enumerated; add SourceFolder for online/offline state monitoring
                        {
                            $SourceFolderArray.Add($FolderItem)
                        }
                    }
                    else  # $SplitOnLeafFolder=$false
                    {
                        $SourceFolderArray.Add($FolderItem)
                    }
                }
            }
        
            if ($SourceFolderArray.Count -gt 0)
            {
                $SourceFolderGroupsJsonHashtable.Add([Int32]$SourceFolderGroupItem.SourceFolderGroup, [array]$SourceFolderArray)
            }
        }
    }

    # Ensure $WallpaperLayoutWeightedListObject GroupNames are unique
    $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObject.Clone()
    foreach ($GroupNameObject in $WallpaperLayoutWeightedListObject)
    {
        if ($GroupNameObject.GroupName -eq 'Default') {continue}  # The 1st instance of Default will win when added to $DataHashtable.WallpaperLayoutWeightedList
        
        $GroupsSameName = $WallpaperLayoutWeightedListObjectCopy.Clone() | Where-Object {$_.GroupName -eq $GroupNameObject.GroupName}

        if ($GroupsSameName.Count -gt 1)
        {
            foreach ($Group in $GroupsSameName)
            {
                # PSCustomObject does not support Clone(), serializing and deserializing to break object linkings
                $GroupSerialized = ConvertTo-Json -InputObject $Group
                $GroupDeserialized = ConvertFrom-Json -InputObject $GroupSerialized
                
                $GroupDeserialized.GroupName = "$(New-Guid)"

                $WallpaperLayoutWeightedListObjectCopy += $GroupDeserialized
            }

            $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObjectCopy | Where-Object {$_.GroupName -ne $GroupNameObject.GroupName}
        }
    }
    $WallpaperLayoutWeightedListObject = $WallpaperLayoutWeightedListObjectCopy.Clone()
    
    # Handle SplitGroupOnSourceFolderGroupSets for WallpaperLayoutWeightedList
    $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObject.Clone()
    foreach ($GroupNameObject in $WallpaperLayoutWeightedListObject)
    {
        if ($GroupNameObject.GroupName -eq 'Default') {continue}  # This process will change GroupName which must be preserved for Default
        
        if (($GroupNameObject.SourceFolderGroupSets.Count -gt 1) -and $GroupNameObject.SplitGroupOnSourceFolderGroupSets)
        {
            $SplitGroup = $WallpaperLayoutWeightedListObjectCopy.Clone() | Where-Object {$_.GroupName -eq $GroupNameObject.GroupName}
            
            foreach ($GroupNameObjectSourceFolderGroupSet in $GroupNameObject.SourceFolderGroupSets)
            {
                # PSCustomObject does not support Clone(), serializing and deserializing to break object linkings
                $GroupSerialized = ConvertTo-Json -InputObject $SplitGroup
                $GroupDeserialized = ConvertFrom-Json -InputObject $GroupSerialized

                $GroupDeserialized.GroupName = "$(New-Guid)"

                # String type supports Clone(), Int32 does not
                $GroupDeserialized.SourceFolderGroupSets = [array]([Int32](([string]$GroupNameObjectSourceFolderGroupSet).Clone()))

                $WallpaperLayoutWeightedListObjectCopy += $GroupDeserialized
            }

            # Remove the original GroupName after the split
            $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObjectCopy | Where-Object {$_.GroupName -ne $GroupNameObject.GroupName}
        }
    }
    $WallpaperLayoutWeightedListObject = $WallpaperLayoutWeightedListObjectCopy.Clone()

    # Handle SplitGroupOnSourceFolderGroupSets for SourceFolderGroupSets
    $SourceFolderGroupSetsObjectCopy = $SourceFolderGroupSetsObject.Clone()
    foreach ($SourceFolderGroupSetItem in $SourceFolderGroupSetsObject)
    {
        $SplitGroupOnSourceFolderGroupSets = [array]($WallpaperLayoutWeightedListObject.Clone() | Where-Object {($_.SplitGroupOnSourceFolderGroupSets) -and ($_.GroupName -ne 'Default') -and ($_.SourceFolderGroupSets -contains $SourceFolderGroupSetItem.SourceFolderGroupSet)})
        
        if (($SourceFolderGroupSetItem.SourceFolderGroups.Count -gt 1) -and ($SplitGroupOnSourceFolderGroupSets.Count -gt 0))
        {
            # Each permutation of [WallpaperLayoutWeightedList with SplitGroupOnSourceFolderGroupSets=$true] x [SourceFolderGroup] will be split out 
            # into new GroupNames with a new SourceFolderGroupSet containing a single SourceFolderGroup.
            # GroupName and SourceFolderGroupSets will have already been split out into GroupNames with a single SourceFolderGroupSet (above).
            foreach ($SplitGroupOnSourceFolderGroupSet in $SplitGroupOnSourceFolderGroupSets)
            {
                foreach ($SourceFolderGroup in $SourceFolderGroupSetItem.SourceFolderGroups)
                {
                    # Create a new SourceFolderGroupSet and assign to a new SplitGroupOnSourceFolderGroupSets GroupName
                    $NewSourceFolderGroupSetNumber = ($SourceFolderGroupSetsObjectCopy.SourceFolderGroupSet | measure -Maximum).Maximum + 1

                    # String type supports Clone(), Int32 does not
                    $SourceFolderGroupsSingleton = [Int32](([string]$SourceFolderGroup).Clone())

                    $NewSourceFolderGroupSet = [PSCustomObject]@{'SourceFolderGroupSet'=[Int32]$NewSourceFolderGroupSetNumber;'SourceFolderGroups'=[array]([Int32]$SourceFolderGroupsSingleton)}

                    $SourceFolderGroupSetsObjectCopy += $NewSourceFolderGroupSet
                    
                    # PSCustomObject does not support Clone(), serializing and deserializing to break object linkings
                    $GroupSerialized = ConvertTo-Json -InputObject $SplitGroupOnSourceFolderGroupSet
                    $GroupDeserialized = ConvertFrom-Json -InputObject $GroupSerialized

                    $GroupDeserialized.GroupName = "$(New-Guid)"

                    $GroupDeserialized.SourceFolderGroupSets = [array]([Int32]$NewSourceFolderGroupSetNumber)

                    $WallpaperLayoutWeightedListObject += $GroupDeserialized
                }

                # Remove the original GroupName after the split
                $WallpaperLayoutWeightedListObject = $WallpaperLayoutWeightedListObject | Where-Object {$_.GroupName -ne $SplitGroupOnSourceFolderGroupSet.GroupName}
            }
        }
    }
    $SourceFolderGroupSetsObject = $SourceFolderGroupSetsObjectCopy.Clone()

    if ($StopwatchInputJsonFileData) {$StopwatchInputJsonFileData.Restart()}
    else {$StopwatchInputJsonFileData = [System.Diagnostics.Stopwatch]::StartNew()}

    # Handle SplitGroupOnSourceFolder for WallpaperLayoutWeightedList
    $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObject.Clone()
    :WallpaperLayoutWeightedListObject foreach ($GroupNameObject in $WallpaperLayoutWeightedListObject)
    {
        if ($GroupNameObject.GroupName -eq 'Default') {continue}  # The "Default" GroupName will always be preserved, and will therefore never be allowed to split

        if (($GroupNameObject.SourceFolderGroupSets.Count -ge 1) -and $GroupNameObject.SplitGroupOnSourceFolder)
        {
            foreach ($SourceFolderGroupSetsInt in $GroupNameObject.SourceFolderGroupSets)
            {
                $SourceFolderGroupSetObject = $SourceFolderGroupSetsObject | Where-Object {$_.SourceFolderGroupSet -eq $SourceFolderGroupSetsInt}

                $SourceFolderObjectArray = @()
                foreach ($SourceFolderGroup in $SourceFolderGroupSetObject.SourceFolderGroups)
                {
                    $SourceFolderGroupsJsonHashtable.([Int32]$SourceFolderGroup) | ForEach-Object {$SourceFolderObjectArray += $_}
                }

                # Create a new 1:1 pair of SourceFolderGroup and SourceFolderGroupSet, and assign SourceFolderGroupSet to a new SplitGroupOnSourceFolder GroupName
                foreach ($SourceFolderObject in $SourceFolderObjectArray)
                {
                    if ([int32]([math]::Round($StopwatchInputJsonFileData.Elapsed.TotalSeconds)) -gt [int32]$DataHashtable.FolderEnumerationTimeLimit) {break WallpaperLayoutWeightedListObject}
                    
                    # PSCustomObject does not support Clone(), serializing and deserializing to break object linkings
                    $SourceFolderSerialized = ConvertTo-Json -InputObject $SourceFolderObject
                    $SourceFolderDeserialized = ConvertFrom-Json -InputObject $SourceFolderSerialized

                    for ($SourceFolderWeight=1; $SourceFolderWeight -le $SourceFolderObject.Weight; $SourceFolderWeight++)
                    {
                        $NewSourceFolderGroupNumber = ($SourceFolderGroupsJsonHashtable.Keys | measure -Maximum).Maximum + 1
                        $NewSourceFolderGroupSetNumber = ($SourceFolderGroupSetsObject.SourceFolderGroupSet | measure -Maximum).Maximum + 1

                        # Add new SourceFolderGroup with a single SourceFolder                
                        $SourceFolderGroupsJsonHashtable.Add([Int32]$NewSourceFolderGroupNumber, $SourceFolderDeserialized)

                        # Add new SourceFolderGroupSet with a single SourceFolderGroup
                        $NewSourceFolderGroupSet = [PSCustomObject]@{'SourceFolderGroupSet'=[Int32]$NewSourceFolderGroupSetNumber;'SourceFolderGroups'=[array]([Int32]$NewSourceFolderGroupNumber)}
                        $SourceFolderGroupSetsObject += $NewSourceFolderGroupSet

                        $GroupSerialized = ConvertTo-Json -InputObject $GroupNameObject
                        $GroupDeserialized = ConvertFrom-Json -InputObject $GroupSerialized
                        
                        $GroupDeserialized.GroupName = "$(New-Guid)"

                        if ($SourceFolderDeserialized.ZoomPercent) {$GroupDeserialized.ZoomPercent = [Int32]$SourceFolderDeserialized.ZoomPercent}
                        if ($SourceFolderDeserialized.RandomFlipHorizontal) {$GroupDeserialized.RandomFlipHorizontal = [System.Convert]::ToBoolean($SourceFolderDeserialized.RandomFlipHorizontal)}

                        # String type supports Clone(), Int32 does not
                        $GroupDeserialized.SourceFolderGroupSets = [array]([Int32](([string]$NewSourceFolderGroupSetNumber).Clone()))

                        # Add new GroupName to WallpaperLayoutWeightedListObject
                        $WallpaperLayoutWeightedListObjectCopy += $GroupDeserialized
                    }
                }
                
                # Remove the original GroupName after the split
                $WallpaperLayoutWeightedListObjectCopy = $WallpaperLayoutWeightedListObjectCopy | Where-Object {$_.GroupName -ne $GroupNameObject.GroupName}
            }
        }
    }
    $WallpaperLayoutWeightedListObject = $WallpaperLayoutWeightedListObjectCopy.Clone()

    $StopwatchInputJsonFileData.Stop()

    # Ingest InputsJsonFile SourceFolderGroupSets
    $SourceFolderGroupJsonHashtable = @{}
    foreach ($SourceFolderGroupSetItem in $SourceFolderGroupSetsObject)
    {
        if ($SourceFolderGroupSetItem.SourceFolderGroups.Count -gt 0)
        {
            $SourceFolderGroupSet_FolderArray = @()
            
            foreach ($SourceFolderGroup in $SourceFolderGroupSetItem.SourceFolderGroups)
            {
                $SourceFolderGroupSet_FolderArray += $SourceFolderGroupsJsonHashtable.([Int32]$SourceFolderGroup)
            }
            
            if (($SourceFolderGroupSet_FolderArray.Weight | measure -Maximum).Maximum -gt 0)
            {
                $SourceFolderGroupSet_FolderHashtable = @{}
                foreach ($SourceFolderGroupSet_FolderItem in $SourceFolderGroupSet_FolderArray)
                {
                    if ((-not $SourceFolderGroupSet_FolderItem.Weight) -or ($SourceFolderGroupSet_FolderItem.Weight -le 0))
                    {continue}
                    
                    if ($SourceFolderGroupSet_FolderHashtable.($SourceFolderGroupSet_FolderItem.Folder))
                    {
                        $SourceFolderGroupSet_FolderHashtable.($SourceFolderGroupSet_FolderItem.Folder) += [Int32]$SourceFolderGroupSet_FolderItem.Weight
                    }
                    else {$SourceFolderGroupSet_FolderHashtable.Add($SourceFolderGroupSet_FolderItem.Folder, [Int32]$SourceFolderGroupSet_FolderItem.Weight)}
                }
            
                $SourceFolderGroupJsonHashtable.Add([Int32]$SourceFolderGroupSetItem.SourceFolderGroupSet, $SourceFolderGroupSet_FolderHashtable)
            }
        }
    }

    # InputsJsonFile WallpaperLayoutWeightedList
    foreach ($GroupNameObject in $WallpaperLayoutWeightedListObject)
    {
        if ([Int32]$GroupNameObject.Weight -gt 0)  # GroupNames in InputsJsonFile.json with a Weight of 0 are effectively disabled
        {
            $DataHashtable.WallpaperLayoutWeightedList.Add($($GroupNameObject.GroupName),@{})
            
            $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('WallpaperLayout', [string]$GroupNameObject.WallpaperLayout)

            $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('Weight', [Int32]$GroupNameObject.Weight)
        }
        else {continue}
        
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('SplitGroupOnSourceFolder', [System.Convert]::ToBoolean($GroupNameObject.SplitGroupOnSourceFolder))
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('IncludeSubdirectories', [System.Convert]::ToBoolean($GroupNameObject.IncludeSubdirectories))
        
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('AlphaNumeric', [System.Convert]::ToBoolean($GroupNameObject.AlphaNumeric))
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('RandomSequence', [System.Convert]::ToBoolean($GroupNameObject.RandomSequence))
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('RandomSequenceTracker', -1)
                
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('ZoomPercent', [Int32]$GroupNameObject.ZoomPercent)
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('RandomFlipHorizontal', [System.Convert]::ToBoolean($GroupNameObject.RandomFlipHorizontal))
        
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('Portrait_Landscape_SurplusBurnoff', [System.Convert]::ToBoolean($GroupNameObject.Portrait_Landscape_SurplusBurnoff))
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('SurplusCandidate_Portrait', [array]$GroupNameObject.SurplusCandidate_Portrait)
        $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('SurplusCandidate_Landscape', [array]$GroupNameObject.SurplusCandidate_Landscape)
        
        if ($GroupNameObject.SourceFolderGroupSets.Count -gt 0)  # GroupNames in InputsJsonFile.json with no SourceFolderGroupSets are effectively disabled
        {
            foreach ($SourceFolderGroupSet in $GroupNameObject.SourceFolderGroupSets)  # Array of SourceFolderGroupSets integers within InputsFileJson
            {
                if (-not $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).SourceFolders)
                {
                    $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).Add('SourceFolders', $SourceFolderGroupJsonHashtable.([Int32]$SourceFolderGroupSet))
                }
                else  # Merge SourceFolders hashtables for more than 1 SourceFolderGroupSet
                {
                    foreach ($SourceFolder in $SourceFolderGroupJsonHashtable.([Int32]$SourceFolderGroupSet).Keys)
                    {
                        if(-not $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).SourceFolders.$SourceFolder)
                        {
                            $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).SourceFolders.Add($SourceFolder, $SourceFolderGroupJsonHashtable.([Int32]$SourceFolderGroupSet).$SourceFolder)  # The value field is the SourceFolder's Weight
                        }
                        else
                        {
                            $DataHashtable.WallpaperLayoutWeightedList.$($GroupNameObject.GroupName).SourceFolders.$SourceFolder += $SourceFolderGroupJsonHashtable.([Int32]$SourceFolderGroupSet).$SourceFolder  # Add Weight to the existing SourceFolder
                        }
                    }
                }
            }
        }
        else {$DataHashtable.WallpaperLayoutWeightedList.Remove($($GroupNameObject.GroupName))}
    }
        
    Set-WallpaperWeightedLayoutList -DataHashtable $DataHashtable
    Get-WallpaperLayoutWeightedState -DataHashtable $DataHashtable
    
    # MaxLayouts will determine the number of Wallpaper image files
    if (($DataHashtable.NumberThreads -gt 1) -and ($PSVersionTable.PSVersion.Major -ge 7))
    {
        $DataHashtable.MaxLayouts = $DataHashtable.NumberThreads
    }
    elseif ($DataHashtable.WallpaperLayoutWeightedList.Default.WallpaperLayout)
    {
        $DataHashtable.MaxLayouts = $DataHashtable.WallpaperLayoutCountHash.($DataHashtable.WallpaperLayoutWeightedList.Default.WallpaperLayout)
    }
    else
    {
        $DataHashtable.MaxLayouts = 4
    }

    $DataHashtable.Add('OutputFolderFastRefresh', 0)

    $DataHashtable.Add('RebuildEverything', $false)

    # $DataHashtable must be returned, Re-initializing $DataHashtable in-function also clears the pointer reference from outside the function
    return $DataHashtable
}


function Test-SourceFolderState {

    # Used to periodically re-check SourceFolder state if 1 or more SourceFolders are offline
    # Report back on state changes
    
    Param(
        [Parameter(Mandatory=$false)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [string]$GroupName,
        [Parameter(Mandatory=$false)]
        [array]$SourceFolderOfflineList
    )
    
    [bool]$SourceFolderOfflineStateChange = $false

    if ($SourceFolderOfflineList)
    {
        [array]$FolderOfflineOldList = $SourceFolderOfflineList
        
        foreach ($SourceFolder in $SourceFolderOfflineList.GetEnumerator())
        {
            if (-not (Test-Path -Path $SourceFolder))
            {
                if ($SourceFolderOfflineList -notcontains $SourceFolder)
                {
                    [array]$SourceFolderOfflineList += $SourceFolder
                }
            }
            elseif ($SourceFolderOfflineList -contains $SourceFolder)
            {
                $SourceFolderOfflineList = [array]($SourceFolderOfflineList | Where-Object {$_ -ne $SourceFolder})
            }
        }

        $FolderOfflineNewList = [array]($SourceFolderOfflineList)
    }

    if ($GroupName)
    {
        $FolderOfflineOldList = [array]($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList)
        
        foreach ($SourceFolder in $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolders.Keys.GetEnumerator())
        {
            if (-not (Test-Path -Path $SourceFolder))
            {
                if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList -notcontains $SourceFolder)
                {
                    [array]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList += $SourceFolder
                }
            }
            elseif ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList -contains $SourceFolder)
            {
                $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList = [array]($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList | Where-Object {$_ -ne $SourceFolder})
            }
        }
        
        $FolderOfflineNewList = [array]($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList)
        
        # Signal $SourceFolder status state changes
        if ((($FolderOfflineOldList | Sort-Object)  -join '') -ne (($FolderOfflineNewList | Sort-Object)  -join '')) {$SourceFolderOfflineStateChange = $true}
    }

    return $SourceFolderOfflineStateChange,[array]$FolderOfflineNewList
}


function Get-SourceFileList {

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [string]$GroupName
    )

    $SourcePictureHash = @{}
    $SourcePictureHash.'SourcePictureList' = [System.Collections.Generic.List[string]]@()
    $SourcePictureHash.'KnownPortrait' = [System.Collections.Generic.List[string]]@()
    $SourcePictureHash.'KnownLandscape' = [System.Collections.Generic.List[string]]@()

    foreach ($SourceFolder in ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolders.Keys | Sort-Object))
    {
        if (-not (Test-Path -Path $SourceFolder))
        {
            [array]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList += $SourceFolder
            
            continue
        }
            
        [Int32]$WeightedMultiplier = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolders.$SourceFolder

        for ($i=1; $i -le $WeightedMultiplier; $i++)
        {
            # Pictures in $SourceFolder and subdirectories, restrict to subdirectories with IncludeRecurseSubdirectoriesTagFile and exclude subdirectories with ExcludeRecurseSubdirectoriesTagFile
            if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.IncludeSubdirectories)  
            {
                # Include recursed SourceFolder directories with IncludeRecurseSubdirectoriesTagFile, then exclude recursed SourceFolder directories with ExcludeRecurseSubdirectoriesTagFile
                if ($DataHashtable.IncludeRecurseSubdirectoriesTagFile -or $DataHashtable.ExcludeRecurseSubdirectoriesTagFile)
                {
                    if ($DataHashtable.IncludeRecurseSubdirectoriesTagFile)
                    {
                        $IncludeRecurseSubdirectoriesTagFileDirectories = [array]((Get-ChildItem -Path "$SourceFolder\*" -Include $DataHashtable.IncludeRecurseSubdirectoriesTagFile -Recurse).DirectoryName | Select-Object -Unique)
                    }
                    
                    if ($DataHashtable.ExcludeRecurseSubdirectoriesTagFile)
                    {
                        $ExcludeRecurseSubdirectoriesTagFileDirectories = [array]((Get-ChildItem -Path "$SourceFolder\*" -Include $DataHashtable.ExcludeRecurseSubdirectoriesTagFile -Recurse).DirectoryName | Select-Object -Unique)
                    }

                    # All directories with Pictures
                    $RecursedPictureDirectories = [System.Collections.Generic.List[string]]((Get-ChildItem -Path "$SourceFolder\*" -Include *.BMP, *.GIF, *.EXIF, *.JPG, *.JPEG, *.PNG, *.TIFF -Recurse).DirectoryName | Select-Object -Unique)

                    # If IncludeRecurseSubdirectoriesTagFile is not found, enumerate all SourceFolder directories with Pictures anyway, but still check for ExcludeRecurseSubdirectoriesTagFile subdirectories
                    if ($IncludeRecurseSubdirectoriesTagFileDirectories.Count -lt 1)
                    {
                        $IncludeRecurseSubdirectoriesTagFileDirectories = $RecursedPictureDirectories | ForEach-Object {$_}
                    }

                    # ExcludeRecurseSubdirectoriesTagFile may be used to prune IncludeRecurseSubdirectoriesTagFile subdirectories
                    # IncludeRecurseSubdirectoriesTagFile directories within ExcludeRecurseSubdirectoriesTagFile directories will not be included
                    
                    $IncludePictureDirectories = [System.Collections.Generic.List[string]]@()
                                        
                    # Picture directories and subdirectories to include
                    foreach ($RecursedPictureDirectory in $RecursedPictureDirectories)
                    {
                        foreach ($IncludeTaggedDirectory in $IncludeRecurseSubdirectoriesTagFileDirectories)
                        {
                            if (($IncludePictureDirectories -notcontains $RecursedPictureDirectory) -and ($RecursedPictureDirectory -like "$IncludeTaggedDirectory\*" -or $RecursedPictureDirectory -eq "$IncludeTaggedDirectory"))
                            {
                                [void]$IncludePictureDirectories.Add($RecursedPictureDirectory)
                            }
                        }
                    }
                                        
                    # Picture directories and subdirectories to exclude
                    $IncludePictureDirListCopy = $IncludePictureDirectories | ForEach-Object {$_} # Loop through a copy, directories cannot be removed from the collection while in the loop
                    
                    foreach ($IncludePictureDir in $IncludePictureDirListCopy)
                    {
                        foreach ($ExcludeTaggedDirectory in $ExcludeRecurseSubdirectoriesTagFileDirectories)
                        {
                            if (($IncludePictureDir -like "$ExcludeTaggedDirectory\*") -or ($IncludePictureDir -eq "$ExcludeTaggedDirectory"))
                            {
                                [void]$IncludePictureDirectories.Remove($IncludePictureDir)
                            }
                        }
                    }

                    foreach ($IncludePictureDirectory in $IncludePictureDirectories)
                    {
                        # The non-recursive method is applied to the final list of all explicitly included/excluded subdirectories
                        (Get-ChildItem -Path "$IncludePictureDirectory\*" -File -Include *.BMP, *.GIF, *.EXIF, *.JPG, *.JPEG, *.PNG, *.TIFF).Name | Where-Object {$_} | ForEach-Object {$SourcePictureHash.SourcePictureList.Add(($IncludePictureDirectory + "\" + $_).Replace('\\',''))}
                    }
                
                }
                else # Pictures in $SourceFolder and all subdirectories
                {
                    (Get-ChildItem -Path "$SourceFolder\*" -Include *.BMP, *.GIF, *.EXIF, *.JPG, *.JPEG, *.PNG, *.TIFF -Recurse).FullName | ForEach-Object {$SourcePictureHash.SourcePictureList.Add($_)}
                }
            }
            else # Pictures in $SourceFolder only, no recursion through subdirectories
            {
                (Get-ChildItem -Path "$SourceFolder\*" -File -Include *.BMP, *.GIF, *.EXIF, *.JPG, *.JPEG, *.PNG, *.TIFF).Name | Where-Object {$_} | ForEach-Object {$SourcePictureHash.SourcePictureList.Add(($SourceFolder + "\" + $_).Replace('\\',''))}
            }

            # If SplitGroupOnSourceFolder=$true, SourceFolder Weight has already been magnified by increasing GroupName count
            if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SplitGroupOnSourceFolder) {break}
        }
    }

    $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash = [hashtable]$SourcePictureHash

    if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count -le 0)
    {
        Get-WallpaperLayoutWeightedState -DataHashtable $DataHashtable
    }

    return
}


function Set-WallpaperWeightedLayoutList {

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable
    )

    if (($DataHashtable.WallpaperLayoutWeightedList.Keys.Count -le 0) -or ($DataHashtable.WallpaperLayoutCountHash.Keys -le 0)) {return}

    [array]$WallpaperLayoutList = $DataHashtable.WallpaperLayoutCountHash.Keys

    $DataHashtable.WallpaperLayoutWeightedRuntimeList = [System.Collections.Generic.List[string]]@()
    
    foreach ($GroupName in $DataHashtable.WallpaperLayoutWeightedList.Keys)
    {
        if ($WallpaperLayoutList -notcontains $DataHashtable.WallpaperLayoutWeightedList.$GroupName.WallpaperLayout)
        {
            continue
        }

        # WallpaperLayoutWeightedRuntimeList
        if ($GroupName -ne 'Default')
        {
            $void = (1..($DataHashtable.WallpaperLayoutWeightedList.$GroupName.Weight)) | ForEach-Object {$DataHashtable.WallpaperLayoutWeightedRuntimeList.Add($GroupName)}
        }
    }

    return
}


function Get-WallpaperLayoutWeightedState {
    
    # Get the functional state of WallpaperLayoutWeightedList: 
    #   1.) GroupName is enabled (Weight > 1)
    #   2.) at least 1 Sourcefolder is specified with a Weight > 0
    #   3.) at least 1 WallpaperLayoutWeightedList SourceFolder is online 
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable
    )

    if ($DataHashtable.WallpaperLayoutWeightedList.Keys.Count -lt 1)
    {
        return
    }

    # Non-Default GroupNames
    $EnabledWeightedGroupNames = [array]($DataHashtable.WallpaperLayoutWeightedList.Keys | Where-Object {$_ -ne 'Default'})
    
    if ($EnabledWeightedGroupNames)
    {
        $EnabledWeightedGroupName_EnabledSourceFolders = @()
        $EnabledWeightedGroupName_EnabledSourceFolders = [array]($EnabledWeightedGroupNames | ForEach-Object {$DataHashtable.WallpaperLayoutWeightedList.$_.SourceFolders})
        $EnabledWeightedGroupName_EnabledSourceFolders = [array]($EnabledWeightedGroupName_EnabledSourceFolders.Keys | Where-Object {$_} | Select-Object -Unique)

        if ($EnabledWeightedGroupName_EnabledSourceFolders.Count -gt 0)
        {
            $DataHashtable.WallpaperLayoutWeightedListEnableState = $true
        
            # Test state of enabled SourceFolders
            $Void_StateChange,$EnabledWeightedGroupName_EnabledSourceFoldersOffline = Test-SourceFolderState -SourceFolderOfflineList $EnabledWeightedGroupName_EnabledSourceFolders

            if ($EnabledWeightedGroupName_EnabledSourceFolders.Count -gt $EnabledWeightedGroupName_EnabledSourceFoldersOffline.Count)
            {
                $DataHashtable.WeightedWallpaperSourceFoldersOnlineState = $true
            }
            else # Use GroupName Default
            {
                $DataHashtable.WeightedWallpaperSourceFoldersOnlineState = $false
            }
        } 
    }
    else 
    {
        $DataHashtable.WallpaperLayoutWeightedListEnableState = $false
        $DataHashtable.WeightedWallpaperSourceFoldersOnlineState = $false
    }

    # Default GroupName only, only used if all other non-Default GroupNames are offline
    if ($DataHashtable.WallpaperLayoutWeightedList.Default.Keys -gt 0)
    {
        $EnabledWeightedGroupName_EnabledSourceFolders = @()
        $EnabledWeightedGroupName_EnabledSourceFolders = [array]($DataHashtable.WallpaperLayoutWeightedList.Default.SourceFolders)
        $EnabledWeightedGroupName_EnabledSourceFolders = [array]($EnabledWeightedGroupName_EnabledSourceFolders.Keys | Where-Object {$_} | Select-Object -Unique)

        if ($EnabledWeightedGroupName_EnabledSourceFolders.Count -gt 0)
        {
            $DataHashtable.DefaultFolderEnableState = $true
        
            # Test state of enabled SourceFolders
            $Void_StateChange,$EnabledWeightedGroupName_EnabledSourceFoldersOffline = Test-SourceFolderState -SourceFolderOfflineList $EnabledWeightedGroupName_EnabledSourceFolders

            if ($EnabledWeightedGroupName_EnabledSourceFolders.Count -gt $EnabledWeightedGroupName_EnabledSourceFoldersOffline.Count)
            {
                $DataHashtable.DefaultSourceFoldersOnlineState = $true
            }
            else # Use GroupName Default
            {
                $DataHashtable.DefaultSourceFoldersOnlineState = $false
            }
        }
    }
    else 
    {
        $DataHashtable.DefaultFolderEnableState = $false
        $DataHashtable.DefaultSourceFoldersOnlineState = $false
    }

    return
}


function Set-RotateFlipFromEXIF {

    # Rotate and/or Flip the image based on EXIF data to match the photographer or artist's intent

    Param(
        [Parameter(Mandatory=$true)]
        [System.Drawing.Image]$SubImageObject
    )

    if ([array]::IndexOf($SubImageObject.PropertyIdList, 274) -eq -1)  # -1 indicates EXIF data was not found
    {
        return $SubImageObject
    }
    else
    {
        switch ($SubImageObject.GetPropertyItem(274).Value[0]) 
        {
            1 {} # Do nothing 
            2 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::RotateNoneFlipX)}
            3 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipNone)}
            4 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate180FlipX)}
            5 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipX)}
            6 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate90FlipNone)}
            7 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipX)}
            8 {$SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::Rotate270FlipNone)}
            default {} # Do nothing 
        }
                
        # $SubImageObject has been modified per EXIF metadata, which should now be removed
        $SubImageObject.RemovePropertyItem(274)
    }

    return $SubImageObject
}


function Get-SubImageObject {

    # All Wallpaper layout functions use Get-SubImageObject in order to determine proper Portrait/Landscape orientation
    # A given picture will be inspected for Portrait/Landscape only once; unused pictures are binned for later use

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [bool]$SelectPortrait
    )

    # Select from pool of KnownPortrait 
    if (($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait.Count -ge 1) -and ($SelectPortrait -eq $true))
    {
        if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.AlphaNumeric) {$SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait[0]}
        elseif ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequence) # Draw from the top of the deck, last added
        {
            [Int32]$SourcePictureListCountMinus1 = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait.Count - 1  # Last item in the list
            
            $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait[$SourcePictureListCountMinus1]
        }
        else {$SubImage = Get-Random -InputObject $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait}

        $SubImageObject = Set-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SubImage $SubImage -Type KnownPortrait
    }
    # Select from pool of KnownLandscape 
    elseif (($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape.Count -ge 1) -and ($SelectPortrait -eq $false))
    {
        if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.AlphaNumeric) {$SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape[0]}
        elseif ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequence) # Draw from the top of the deck, last added
        {
            [Int32]$SourcePictureListCountMinus1 = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape.Count - 1  # Last item in the list

            $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape[$SourcePictureListCountMinus1]
        }
        else {$SubImage = Get-Random -InputObject $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape}

        $SubImageObject = Set-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SubImage $SubImage -Type KnownLandscape
    }
    else
    {
        if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.Portrait_Landscape_SurplusBurnoff -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Portrait.Count -gt 0))
        {
            $PortraitSurplusBurnoff = $true
        }
        else {$PortraitSurplusBurnoff = $false}

        if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.Portrait_Landscape_SurplusBurnoff -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Landscape.Count -gt 0))
        {
            $LandscapeSurplusBurnoff = $true
        }
        else {$LandscapeSurplusBurnoff = $false}

        do 
        {
            if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count -lt 1)
            {
                $DataHashtable.GetSubImageObjectSuccess = $false
                
                return
            }
            
            if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.AlphaNumeric) # AlphaNumeric, always draw and remove from the bottom of the deck
            {
                $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList[0]
            }
            elseif ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequence -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker -le -1))  # Start RandomSequence
            {
                $i = Get-Random -Minimum 0 -Maximum $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count

                $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList[$i]

                $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker = $i
            }
            elseif ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequence -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker -ge 0))  # Continue RandomSequence
            {
                if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker -lt $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count)
                {
                    # Continue sequence from where we left off
                    $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList[$DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker]
                }
                else  
                {
                    # Draw from the top, shift the sequence backwards
                    [Int32]$SourcePictureListCountMinus1 = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count - 1

                    $SubImage = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList[$SourcePictureListCountMinus1]
                }
            }
            else  # Draw randomly
            {
                $SubImage = (Get-Random -InputObject $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList)
            }

            $SubImageObject = Set-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SubImage $SubImage -Type SourcePictureList

            # Portrait
            if ($SubImageObject.Width -le $SubImageObject.Height)
            {
                $IsPortrait = $true
            }
            else #Landscape
            {
                $IsPortrait = $false
            }

            if (($IsPortrait -and $SelectPortrait) -or ((-not $IsPortrait) -and (-not $SelectPortrait)))
            {
                $SubImageFound = $true

                [void]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Remove($SubImage)
            }
            elseif ($IsPortrait) # The Layout function is asking for a Landscape, add found Portrait for later.
            {
                $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait.Add($SubImage)
                [void]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Remove($SubImage)

                $SubImageFound = $false

                # Dispose if $SubImageObject is [System.Drawing.Bitmap], else remove the variable
                if ($SubImageObject -and (($SubImageObject.GetType().FullName -eq 'System.Drawing.Bitmap') -and ($SubImageObject.Width -ge 1) -and ($SubImageObject.Height -ge 1)))
                {
                    $SubImageObject.Dispose()
                } 
                else {Remove-Variable SubImageObject}

                # Return for PortraitLandscapeSurplusBurnoff, in case a Portrait heavy SourceFolder is being applied to a Landscape Wallpaper Layout
                if ($PortraitSurplusBurnoff -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait.Count -gt 50))
                {
                    $DataHashtable.GetSubImageObjectSuccess = $false

                    return
                }
            }
            else # The Layout function is asking for a Portrait, add found Landscape for later.
            {
                $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape.Add($SubImage)
                [void]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Remove($SubImage)

                $SubImageFound = $false

                # Dispose if $SubImageObject is [System.Drawing.Bitmap], else remove the variable
                if ($SubImageObject -and (($SubImageObject.GetType().FullName -eq 'System.Drawing.Bitmap') -and ($SubImageObject.Width -ge 1) -and ($SubImageObject.Height -ge 1)))
                {
                    $SubImageObject.Dispose()
                } 
                else {Remove-Variable SubImageObject}

                # Return for PortraitLandscapeSurplusBurnoff, in case a Landscape heavy SourceFolder is being applied to a Portrait Wallpaper Layout
                if ($LandscapeSurplusBurnoff -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape.Count -gt 50))
                {
                    $DataHashtable.GetSubImageObjectSuccess = $false

                    return
                }
            }
        }
        until ($SubImageFound -or ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count -le 1))
    }

    if ($SubImageObject -and (($SubImageObject.GetType().FullName -ne 'System.Drawing.Bitmap') -or ($SubImageObject.Width -le 0) -or ($SubImageObject.Height -le 0)))
    {
        $DataHashtable.GetSubImageObjectSuccess = $false

        Remove-Variable SubImageObject
    }
    
    return $SubImageObject
}


function Set-SubImageObject {
    
    # Create SubImageObject using the full file name and EXIF data
        
    Param(
    [Parameter(Mandatory=$true)]
    [hashtable]$DataHashtable,
    [Parameter(Mandatory=$true)]
    [string]$GroupName,
    [Parameter(Mandatory=$true)]
    [string]$SubImage,
    [Parameter(Mandatory=$true)]
    [ValidateSet('KnownPortrait','KnownLandscape','SourcePictureList')]
    [string]$Type
    )

    if (-not [System.IO.File]::Exists($SubImage))
    {
        $DataHashtable.RebuildEverything = $true
            
        $DataHashtable.GetSubImageObjectSuccess = $false
            
        return
    }

    # $SubImageObject will be disposed in Set-SubImage
    $SubImageObject = new-object System.Drawing.Bitmap $SubImage

    if ($SubImageObject.GetType().FullName -eq 'System.Drawing.Bitmap')
    {
        # System.Drawing.Graphics does not support Indexed color PixelFormats, images using them will need to be converted to RGB
        if ($SubImageObject.PixelFormat -like "*Indexed")
        {
            $RgbSubImageObject = New-Object System.Drawing.Bitmap -ArgumentList $SubImageObject.Width, $SubImageObject.Height, Format32bppRgb

            $RgbNonIndexedImage = [System.Drawing.Graphics]::FromImage($RgbSubImageObject)
            $RgbNonIndexedImage.DrawImage($SubImageObject, 0, 0, $SubImageObject.Width, $SubImageObject.Height)

            $SubImageObject = $RgbSubImageObject.Clone()

            $RgbNonIndexedImage.Dispose()
        }
        
        # Some Pictures have embedded EXIF data for rotate and flip, interpretted by Windows, but not immediately interpretted by System.Drawing.Bitmap
        $SubImageObject = Set-RotateFlipFromEXIF -SubImageObject $SubImageObject

        [void]$DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.$Type.Remove($SubImage)
    }
    else
    {
        Remove-Variable SubImageObject
            
        return
    }

    return $SubImageObject
}


function Set-SubImage {

    Param(
        [Parameter(Mandatory=$false)]
        [System.Drawing.Image]$SubImageObject,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Image]$WallpaperObject,
        [Parameter(Mandatory=$true)]
        [Int32]$OriginPointX,
        [Parameter(Mandatory=$true)]
        [Int32]$OriginPointY,
        [Parameter(Mandatory=$true)]
        [Int32]$SubImageWidth,
        [Parameter(Mandatory=$true)]
        [Int32]$SubImageHeight,
        [Parameter(Mandatory=$true)]
        [double]$AspectRatio,
        [Parameter(Mandatory=$false)]
        [Int32]$ZoomPercent,
        [Parameter(Mandatory=$false)]
        [bool]$RandomFlipHorizontal
    )
    
    # Poison pill the Wallpaper creation process here, rather than in the many Wallpaper layout functions and Set-SubImage function calls
    if ((-not $SubImageObject) -or ($SubImageObject.GetType().FullName -ne 'System.Drawing.Bitmap') -or ($SubImageObject.Width -le 0) -or ($SubImageObject.Height -le 0))
    {
        $DataHashtable.GetSubImageObjectSuccess = $false

        if ($SubImageObject) {Remove-Variable SubImageObject ; return $WallpaperObject}
    }
    
    if (($RandomFlipHorizontal -eq $true) -and (((0,1) | Get-Random) -eq 1)) 
    {
        $SubImageObject.RotateFlip([System.Drawing.RotateFlipType]::RotateNoneFlipX)
    }

    if (($SubImageObject.Width / $SubImageObject.Height) -ne $AspectRatio)
    {
        $SubImageObject = Set-CropImage -SubImageObject $SubImageObject -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent
    } 
   
    # Draw resized $SubImage and draw within $WallpaperObject
    $DrawSubimageInWallpaper = [System.Drawing.Graphics]::FromImage($WallpaperObject)
    $DrawSubimageInWallpaper.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $DrawSubimageInWallpaper.InterpolationMode = [System.Drawing.Drawing2D.QualityMode]::High
    $DrawSubimageInWallpaper.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $DrawSubimageInWallpaper.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $DrawSubimageInWallpaper.DrawImage($SubImageObject, $OriginPointX, $OriginPointY, $SubImageWidth, $SubImageHeight)
                
    $DrawSubimageInWallpaper.Dispose()
    $SubImageObject.Dispose()

    return $WallpaperObject
}


function Set-CropImage {
    # Crop image based on desired aspect ratio

    Param(
        [Parameter(Mandatory=$true)]
        [System.Drawing.Image]$SubImageObject,
        [Parameter(Mandatory=$true)]
        [double]$AspectRatio,
        [Parameter(Mandatory=$false)]
        [Int32]$ZoomPercent
    )

    if ((-not $ZoomPercent) -or ($ZoomPercent -lt 100)) {$ZoomPercent = 100}
    
    if ($SubImageObject.GetType().FullName -ne 'System.Drawing.Bitmap')
    {
        $DataHashtable.GetSubImageObjectSuccess = $false

        if ($SubImageObject) {Remove-Variable SubImageObject ; return}
    }
    
    if (($SubImageObject.Width / $SubImageObject.Height) -gt $AspectRatio) # Crop vertically
    {
        $CropWidth = [Math]::Round($SubImageObject.Height * $AspectRatio * (100/$ZoomPercent))  # AspectRatio = Width/Height; Height * (Width/Height) = Width
        $CropHeight = [Math]::Round($SubImageObject.Height * (100/$ZoomPercent))

        # Do not use $CropWidth or $CropHeight, to prevent compounding rounding errors
        $OriginPointX = [Math]::Round(($SubImageObject.Width - ($SubImageObject.Height * $AspectRatio * (100/$ZoomPercent)))/2)
        $OriginPointY = 0 + [Math]::Round(($SubImageObject.Height - ($SubImageObject.Height * (100/$ZoomPercent)))/2)
    }
    else # Crop horizontally
    {
        $CropWidth = [Math]::Round($SubImageObject.Width * (100/$ZoomPercent))
        $CropHeight = [Math]::Round(($SubImageObject.Width / $AspectRatio) * (100/$ZoomPercent))  # AspectRatio = Width/Height; Width / (Width/Height) = Height

        # Do not use $CropWidth or $CropHeight, to prevent compounding rounding errors
        $OriginPointX = 0 + [Math]::Round(($SubImageObject.Width - ($SubImageObject.Width * (100/$ZoomPercent)))/2)
        $OriginPointY = [Math]::Round(($SubImageObject.Height - (($SubImageObject.Width / $AspectRatio) * (100/$ZoomPercent)))/2)
    }

    $CropRectangle = new-object Drawing.Rectangle $OriginPointX, $OriginPointY, $CropWidth, $CropHeight
    $SubImageObjectRectangle = new-object Drawing.Rectangle 0, 0, $SubImageObject.Width, $SubImageObject.Height

    $Units = [System.Drawing.GraphicsUnit]::Pixel
    $CroppedImage = [System.Drawing.Graphics]::FromImage($SubImageObject)
    $CroppedImage.DrawImage($SubImageObject, $SubImageObjectRectangle, $CropRectangle, $Units)

    $CroppedImage.Dispose()

    return $SubImageObject
}


function New-SurplusWallpaperImage {

# Use for mixed portrait and landscape layouts.  Surplus portrait or landscape pictures will be consumed by all-portrait or all-landscape layouts, respectively.
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$false)]
        [System.Object]$Stopwatch
    )

    if (($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownPortrait.Count -ge 30) -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Portrait.Count -ge 1))
    {
        $GroupNameSurplusCandidate = Get-Random $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Portrait
        $SurplusWallpaperFunction = "New-Wallpaper_$GroupNameSurplusCandidate"
    }
    elseif (($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.KnownLandscape.Count -ge 30) -and ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Landscape.Count -ge 1))
    {
        $GroupNameSurplusCandidate = Get-Random $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SurplusCandidate_Landscape
        $SurplusWallpaperFunction = "New-Wallpaper_$GroupNameSurplusCandidate"
    }

    if ($SurplusWallpaperFunction)
    {
        $Stopwatch = New-WallpaperImage -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperFunction $SurplusWallpaperFunction -Stopwatch $Stopwatch

        $SurplusWallpaperCreated = $true
    }
    else
    {
        $SurplusWallpaperCreated = $false
    }
    
    return $SurplusWallpaperCreated,$Stopwatch
}


function Test-MultiThread {

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [System.Object]$Stopwatch
    )

    # Switch to multi-thread if Wallpaper creation is exceeding both TimeBetweenCreation and 10 seconds and AutoMultiThread is enabled
    if ($DataHashtable.AutoMultiThread -and (-not $DataHashtable.Thread) -and $Stopwatch -and ($DataHashtable.NumberThreads -lt 2) -and (-not $DataHashtable.AutoMultiThreadActive) -and ([int32]([math]::Round($Stopwatch.Elapsed.TotalSeconds)) -gt [int32](@($DataHashtable.TimeBetweenCreation,10) | measure -Maximum).Maximum))
    {
        $DataHashtable.NumberThreads = [Int32](([string]$DataHashtable.MaxLayouts).Clone())
        
        # If ThreadJobs are started, sleep within Start-ThreadThreadJobs until or more threads are terminated, then continue here and exit Test-MultiThread
        $MultiThreadActivated = Start-ThreadThreadJobs -DataHashtable $DataHashtable -AutoMultiThreadActive True  # AutoMultiThreadActive is a string, Start-ThreadJob and $using: fails with boolean

        $DataHashtable.NumberThreads = 1
    }
    elseif ($Stopwatch -and $DataHashtable.AutoMultiThreadActive -and $DataHashtable.Thread -and ([int32]([math]::Round($Stopwatch.Elapsed.TotalSeconds)) -lt [int32](@($DataHashtable.TimeBetweenCreation,10) | measure -Maximum).Maximum))
    {
        exit  # Terminate the ThreadJob process, return to Start-ThreadThreadJobs function call within this function
    }
    else
    {
        $MultiThreadActivated = $false
    }

    return $MultiThreadActivated
}


function Test-UpdateRequested {

    # Test for requested updates to InputsJsonFile.json or rebuild of $DataHashtable

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [System.Object]$Stopwatch
    )
            
    if ($DataHashtable.RebuildEverything)
    {
        $UpdateRequested = $true
    }
    elseif ((Get-ChildItem -Path $DataHashtable.OutputFolder).Count -ge 1)
    {
        # A negative or relatively small positive value of JsonInputFileSecondsLastUpdate indicates either 1.) InputsJsonFile has been recently updated, or 2.) Wallpaper image creation if failing after InputsJsonFile has been saved
        $JsonInputFileSecondsLastUpdate = (New-TimeSpan -Start (Get-Item $DataHashtable.InputsJsonFile).LastWriteTime -End (((Get-ChildItem -Path $DataHashtable.OutputFolder).LastWriteTime | Sort-Object -Descending)[0])).TotalSeconds

        # Re-evaluate $InputsJsonFile if the time since its last update is less than the time between TimeBetweenCreation and the last Wallpaper file creation time
        if ($JsonInputFileSecondsLastUpdate -lt ($DataHashtable.TimeBetweenCreation + [math]::Round($Stopwatch.Elapsed.TotalSeconds)))
        {
            # Terminate the ThreadJob process and return to Start-ThreadThreadJobs function
            # Unhandled Exception Case:  Thread tear-down sawtoothing will occur if all images are surplus portrait or landscape, we will error on the side of thread tear-down 
            if ($DataHashtable.Thread)
            {
                exit
            }

            $UpdateRequested = $true
        }
    }
    else
    {
        $UpdateRequested = $false
    }

    return $UpdateRequested
}


function New-WallpaperImage {

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [string]$WallpaperFunction,
        [Parameter(Mandatory=$false)]
        [System.Object]$Stopwatch
    )

    # Track the amount of time to create Wallpaper images, then subtract from TimeBetweenCreation sleep time
    if ($Stopwatch) {$Stopwatch.Restart()}
    else {$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()}

    # Wallpaper Image creation will a lot of CPU resources, pause creation if a PauseOnProcess process is running
    if (($DataHashtable.PauseOnProcess | ForEach-Object {(Get-Process).ProcessName -contains $_} | Select-Object -Unique))
    {
        $Stopwatch.Stop()

        Start-Sleep -Seconds $DataHashtable.TimeBetweenCreation

        return $Stopwatch
    }
    
    $WallpaperObject = new-object System.Drawing.Bitmap $DataHashtable.WallpaperDimensions.WallpaperWidth,$DataHashtable.WallpaperDimensions.WallpaperHeight
    
    # Initialize GetSubImageObjectSuccess = $true.  If there is a failure of Get-SubImageObject in a WallpaperLayout function, GetSubImageObjectSuccess will be set to false.
    # Saved Picture names will not be incremented until Save-WallpaperImage is called by this function, the next do loop operation in the main body is a retry.
    $DataHashtable.GetSubImageObjectSuccess = $true
    
    # Create Wallpaper Layout Image
    $WallpaperObject = & $WallpaperFunction -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperObject $WallpaperObject
    
    if ($DataHashtable.GetSubImageObjectSuccess)
    {
        Save-WallpaperImage -WallpaperObject $WallpaperObject -DataHashtable $DataHashtable -GroupName $GroupName

        $Stopwatch.Stop()

        # Sleep for the amount of time left over when subtracting the Stopwatch time from TimeBetweenCreation
        if ($DataHashtable.OutputFolderFastRefresh -gt 0) # Ignore TimeBetweenCreation if InputsJsonFile.json has been changed, until all Wallpapers in OutputFolder have been refreshed
        {
            $DataHashtable.OutputFolderFastRefresh--

            $SleepTime = 0
        }
        elseif ($DataHashtable.Thread -gt 0)  # Sleep longer based on NumberThreads creating Wallpapers in parallel
        {
            $SleepTime = [Int32](($DataHashtable.TimeBetweenCreation * $DataHashtable.NumberThreads) - [math]::Round($Stopwatch.Elapsed.TotalSeconds))
        }
        else  # Standard TimeBetweenCreation, including PictureNameOverride
        {
            $SleepTime = [Int32]($DataHashtable.TimeBetweenCreation - [math]::Round($Stopwatch.Elapsed.TotalSeconds))
        }

        if ($SleepTime -gt 0)
        {
            Start-Sleep -Seconds $SleepTime
        }
    }
    else
    {
        $Stopwatch.Stop()
    }

    return $Stopwatch
}


function Save-WallpaperImage {

    Param(
        [Parameter(Mandatory=$false)]
        [System.Drawing.Bitmap]$WallpaperObject,
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName
    )

    
    function Save-File {

        # Windows streams images without a file lock, Save-File will prevent tearing of Wallpaper images and black desktops
        
        Param(
            [Parameter(Mandatory=$true)]
            [System.Drawing.Bitmap]$WallpaperObject,
            [Parameter(Mandatory=$true)]
            [string]$FilePath
        )

        $ImageFormat = $WallpaperObject.RawFormat

        $WallpaperObject.Save("$($FilePath)_temp",$ImageFormat)

        if ([math]::Round(((Get-Date) - [System.IO.File]::GetLastAccessTime($FilePath)).TotalMilliseconds) -lt 1000)
        {
            sleep -Seconds 1
        }

        if ([System.IO.File]::Exists($FilePath)) {Remove-Item -Path $FilePath}
        Rename-Item -Path "$($FilePath)_temp" -NewName $FilePath -Force

        $WallpaperObject.Dispose()
        
        return
    }
    
    
    # Do not increment PictureName counters if $WallpaperObject is not a valid [System.Drawing.Bitmap] object due to possible corruption by corrupted Picture files; if corrupt, silently return to main body Do loop
    if ($WallpaperObject -and ($WallpaperObject.GetType().FullName -eq 'System.Drawing.Bitmap'))
    {
        if ($DataHashtable.Thread -and ($DataHashtable.Thread -gt 0) -and ($DataHashtable.Thread.GetType().Name -eq 'Int32'))
        {
            Save-File -WallpaperObject $WallpaperObject -FilePath "$($DataHashtable.OutputFolder)\$($DataHashtable.Thread).png"
        }
        elseif ($DataHashtable.PictureNameOverride -and ($DataHashtable.PictureNameOverride -gt 0) -and ($DataHashtable.PictureNameOverride.GetType().Name -eq 'Int32'))
        {
            Save-File -WallpaperObject $WallpaperObject -FilePath "$($DataHashtable.OutputFolder)\$($DataHashtable.PictureNameOverride).png"

            $DataHashtable.PictureNameOverride++
            if ($DataHashtable.PictureNameOverride -gt $DataHashtable.MaxLayouts) {$DataHashtable.PictureNameOverride = 1}
        }
        elseif ($DataHashtable.LayoutTracker -and ($DataHashtable.LayoutTracker -gt 0) -and ($DataHashtable.LayoutTracker.GetType().Name -eq 'Int32'))
        {
            Save-File -WallpaperObject $WallpaperObject -FilePath "$($DataHashtable.OutputFolder)\$($DataHashtable.LayoutTracker).png"

            $DataHashtable.LayoutTracker++
            if ($DataHashtable.LayoutTracker -gt $DataHashtable.WallpaperLayoutCountHash.($DataHashtable.WallpaperLayoutWeightedList.$GroupName.WallpaperLayout)) {$DataHashtable.LayoutTracker = 1}
        }

        $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomSequenceTracker = -1
    }

    return
}


function Start-ThreadThreadJobs {

    # Start-ThreadJob is only supported in PowerShell version 7+

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$false)]
        [string]$AutoMultiThreadActive  # AutoMultiThreadActive is a string, Start-ThreadJob and $using: fails with boolean
    )

    if (($DataHashtable.NumberThreads -gt 1) -and ($PSVersionTable.PSVersion.Major -ge 7))
    {
        $ScriptFile = $PSCommandPath
        $InputsJsonFile = $DataHashtable.InputsJsonFile

        for ($Thread = 1; $Thread -le $DataHashtable.NumberThreads; $Thread++)
        {
            Start-ThreadJob -ScriptBlock {pwsh.exe -File $using:ScriptFile -InputsJsonFile $using:InputsJsonFile -Thread $using:Thread -AutoMultiThreadActive $using:AutoMultiThreadActive} -ThrottleLimit 6
        }

        $ThreadJobsStarted = $true
    }
    else
    {
        $ThreadJobsStarted = $false
    }

    # Sleep here repeadedly until 1 or more ThreadJobs is terminated
    do {sleep -Seconds $DataHashtable.TimeBetweenCreation} while ((Get-Job -State Running).Count -eq $DataHashtable.NumberThreads)

    Get-Job | Remove-Job -Force

    do 
    {
        Get-WmiObject Win32_Process | Where-Object {($_.ProcessName -eq 'pwsh.exe') -and ($_.ParentProcessId -eq $PID)} | ForEach-Object {Stop-Process -Id $_.ProcessId -Force}
    }
    while ((Get-WmiObject Win32_Process | Where-Object {($_.ProcessName -eq 'pwsh.exe') -and ($_.ParentProcessId -eq $PID)}).Count -ge 1)

    return $ThreadJobsStarted
}


function New-Wallpaper_Landscape_1_Base6Thirds {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }

    for ($i=1; $i -le 6; $i++) 
    {
        switch ($Layout)
        {
            1 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 2560 ; $SubImageHeight = 1440 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    2 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    default {break}
                }
            }
            2 {            
                switch ($i)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 2560 ; $SubImageHeight = 1440 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    2 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    4 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    default {break}
                }
            }
            3 {            
                switch ($i)
                {
                    1 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $SubImageWidth = 2560 ; $SubImageHeight = 1440 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    4 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    6 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    default {break}
                }
                }
            4 {            
                switch ($i)
                {
                    1 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 2560 ; $SubImageHeight = 1440 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    3 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    4 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    6 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 16/9 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    default {break}
                }
            }
        }
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
        
        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        # Add SubImageObject to WallpaperObject, after cropping to acheive the AspectRatio and applying RandomFlipHorizontal
        $WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal
    }

    return $WallpaperObject
}


function New-Wallpaper_Landscape_2_Elements_Thirds {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )


    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$true)]
            [ValidateSet('SmallBlock','LargeBlock','LongBlock')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$Layout,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
                        
        switch ($SubLayoutType)
        {
            'SmallBlock' {
        
                $SmallBlockLayout = Get-Random -Minimum 1 -Maximum 3   
                
                for ($SubLayoutNum=1; $SubLayoutNum -le 3; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
                {
                    switch ($SmallBlockLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 540 ; $SubImageHeight = 360 ; $AspectRatio = 540/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 540 ; $OriginPointElementY = 0 ;$SubImageWidth = 1100 ; $SubImageHeight = 720 ; $AspectRatio = 1100/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 540 ; $SubImageHeight = 360 ; $AspectRatio = 540/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1100 ; $SubImageHeight = 720 ; $AspectRatio = 1100/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 1100 ; $OriginPointElementY = 0 ; $SubImageWidth = 540 ; $SubImageHeight = 360 ; $AspectRatio = 540/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1100 ; $OriginPointElementY = 360 ; $SubImageWidth = 540 ; $SubImageHeight = 360 ; $AspectRatio = 540/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
            }
            'LargeBlock' {

                $LargeBlockLayout = Get-Random -Minimum 1 -Maximum 5
                
                for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
                {
                    switch ($LargeBlockLayout)
                    {
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 1466 ; $SubImageHeight = 960 ; $AspectRatio = 1466/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 733 ; $OriginPointElementY = 0 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1466 ; $OriginPointElementY = 0 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 1466 ; $OriginPointElementY = 480 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 1466 ; $OriginPointElementY = 960 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1466 ; $SubImageHeight = 960 ; $AspectRatio = 1466/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 733 ; $OriginPointElementY = 960 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1466 ; $OriginPointElementY = 0 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 1466 ; $OriginPointElementY = 480 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 1466 ; $OriginPointElementY = 960 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 734 ; $OriginPointElementY = 0 ; $SubImageWidth = 1466 ; $SubImageHeight = 960 ; $AspectRatio = 1466/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 734 ; $OriginPointElementY = 960 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1467 ; $OriginPointElementY = 960 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        4 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 734 ; $OriginPointElementY = 480 ; $SubImageWidth = 1466 ; $SubImageHeight = 960 ; $AspectRatio = 1466/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 734 ; $OriginPointElementY = 0 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1467 ; $OriginPointElementY = 0 ; $SubImageWidth = 733 ; $SubImageHeight = 480 ; $AspectRatio = 733/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 734 ; $SubImageHeight = 480 ; $AspectRatio = 734/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
            }
            'LongBlock' {
        
                for ($SubLayoutNum=1; $SubLayoutNum -le 2; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1100 ; $SubImageHeight = 720 ; $AspectRatio = 1100/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 1100 ; $OriginPointElementY = 0 ; $SubImageWidth = 1100 ; $SubImageHeight = 720 ; $AspectRatio = 1100/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function New-Wallpaper_Landscape_2_Elements_Thirds body ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
        
    for ($SubLayoutNum=1; $SubLayoutNum -le 5; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            1 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LargeBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 2200 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2200 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2200 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LongBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            2 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LargeBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 2200 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2200 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2200 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LongBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            3 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 1640 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LargeBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1640 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LongBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            4 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 1640 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LargeBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType SmallBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1640 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayoutType LongBlock -Layout $Layout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        }

        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Landscape_3_Elements_Base3 {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )


    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('Landscape')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
                
        if ($OriginPointX -eq 0)
        {
            $DataHashtable.'RandomLayoutTracker' = @(1,2,3)
        }
        
        $SubLayout = Get-Random -InputObject $DataHashtable.RandomLayoutTracker
        $DataHashtable.RandomLayoutTracker = $DataHashtable.RandomLayoutTracker | Where-Object {$_ -ne $SubLayout}

        for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutType)
            {
                'Landscape' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 640 ; $OriginPointElementY = 0 ;$SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 640 ; $OriginPointElementY = 360 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 640 ; $OriginPointElementY = 720 ;$SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 640 ; $OriginPointElementY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 640 ; $OriginPointElementY = 1440 ;$SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 1800 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 640 ; $OriginPointElementY = 1800 ; $SubImageWidth = 640 ; $SubImageHeight = 360 ; $AspectRatio = 640/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function Landscape_3_Elements_Base ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 3; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            { @(1..6) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Portrait_1_2x2_3x3 {

    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($i=1; $i -le 14; $i++) 
    {
        switch ($Layout) # Intended for 3840 pixels wide x 2160 pixels high (standard 4K)
        {
            1 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    4 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    5 {$OriginPointX = 1707 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 1707 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    7 {$OriginPointX = 2347 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 2347 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 2347 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    12 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 3413 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    10 {$OriginPointX = 2773 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2773 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    default {break}
                }
            }
            2 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    5 {$OriginPointX = 1707 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 1707 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    7 {$OriginPointX = 2347 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 2347 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 2347 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    10 {$OriginPointX = 2773 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2773 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    12 {$OriginPointX = 2773 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    13 {$OriginPointX = 3200 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3200 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }            
            3 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 427 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    6 {$OriginPointX = 1067 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                                        
                    7 {$OriginPointX = 2347 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 2347 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 2347 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    10 {$OriginPointX = 2773 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2773 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    12 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 3413 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }               
            4 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 427 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    7 {$OriginPointX = 1067 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 1067 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 1067 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    6 {$OriginPointX = 1493 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    10 {$OriginPointX = 2773 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2773 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    12 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 3413 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }             
            5 {            
                switch ($i)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 0 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    3 {$OriginPointX = 640 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    4 {$OriginPointX = 640 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 640 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    6 {$OriginPointX = 1067 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    7 {$OriginPointX = 1067 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 1067 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    9 {$OriginPointX = 1493 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 1493 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    11 {$OriginPointX = 2133 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    12 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 3413 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }              
            6 {            
                switch ($i)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 427 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    6 {$OriginPointX = 1067 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    7 {$OriginPointX = 1067 ; $OriginPointY = 720 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    8 {$OriginPointX = 1067 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    9 {$OriginPointX = 1493 ; $OriginPointY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 1493 ; $OriginPointY = 1080 ; $SubImageWidth = 640 ; $SubImageHeight = 1080 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    11 {$OriginPointX = 2133 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    12 {$OriginPointX = 2133 ; $OriginPointY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 2133 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    14 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 2160 ; $AspectRatio = 1280/2160 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }
        }
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal

        # Add SubImageObject to WallpaperObject, after cropping to acheive the AspectRatio and applying RandomFlipHorizontal
        $WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal
    }

    return $WallpaperObject
}


function New-Wallpaper_Portrait_2_NarrowThirds {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    
    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$true)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('Small','Large')]
            [string]$SubLayoutSize,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        for ($SubLayoutNum=1; $SubLayoutNum -le 4; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutSize)
            {
                'Small' 
                {
                    switch ($SubLayout)
                    {  
                        { @(1..6) -contains $_ } {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 320 ; $SubImageHeight = 720 ; $AspectRatio = 320/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 320 ; $OriginPointElementY = 0 ;$SubImageWidth = 320 ; $SubImageHeight = 720 ; $AspectRatio = 320/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 640 ; $OriginPointElementY = 0 ; $SubImageWidth = 320 ; $SubImageHeight = 720 ; $AspectRatio = 320/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 960 ; $OriginPointElementY = 0 ; $SubImageWidth = 320 ; $SubImageHeight = 720 ; $AspectRatio = 320/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'Large'
                {
                    switch ($SubLayout)
                    {  
                        { @(1..6) -contains $_ } {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1440 ; $AspectRatio = 640/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 640 ; $OriginPointElementY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1440 ; $AspectRatio = 640/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1280 ; $OriginPointElementY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1440 ; $AspectRatio = 640/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1920 ; $OriginPointElementY = 0 ; $SubImageWidth = 640 ; $SubImageHeight = 1440 ; $AspectRatio = 640/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    } 
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function New-Wallpaper_Portrait_2_NarrowThirds body ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
        
    $SubLayoutList = @(1..6) # Permutations within function Set-SubLayout

    for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        $SubLayout = Get-Random -InputObject $SubLayoutList

        $SubLayoutList = $SubLayoutList | Where-Object {$_ -ne $SubLayout}
        
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            1 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            2 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            3 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            4 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Portrait_3_Elements_Base3_Narrow {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )


    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('PortraitWide','PortraitNarrow')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        if ($SubLayoutType -eq 'PortraitWide') {$SubLayout = Get-Random -Minimum 1 -Maximum 5}
        if ($SubLayoutType -eq 'PortraitNarrow') {$SubLayout = Get-Random -Minimum 1 -Maximum 5}

        for ($SubLayoutNum=1; $SubLayoutNum -le 9; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutType)
            {
                'PortraitWide' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 900 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 450 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 450 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        4 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 450 ; $OriginPointElementY = 720 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'PortraitNarrow' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 912 ; $SubImageHeight = 1440 ; $AspectRatio = 912/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ;$SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 456 ; $OriginPointElementY = 0 ; $SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 912 ; $OriginPointElementY = 0 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 912 ; $OriginPointElementY = 360 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 912 ; $OriginPointElementY = 720 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                7 {$OriginPointElementX = 912 ; $OriginPointElementY = 1080 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                8 {$OriginPointElementX = 912 ; $OriginPointElementY = 1440 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                9 {$OriginPointElementX = 912 ; $OriginPointElementY = 1800 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ;$SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 456 ; $OriginPointElementY = 1440 ; $SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 912 ; $SubImageHeight = 1440 ; $AspectRatio = 912/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 912 ; $OriginPointElementY = 0 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 912 ; $OriginPointElementY = 360 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 912 ; $OriginPointElementY = 720 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                7 {$OriginPointElementX = 912 ; $OriginPointElementY = 1080 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                8 {$OriginPointElementX = 912 ; $OriginPointElementY = 1440 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                9 {$OriginPointElementX = 912 ; $OriginPointElementY = 1800 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1080 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 1800 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                7 {$OriginPointElementX = 228 ; $OriginPointElementY = 0 ; $SubImageWidth = 912 ; $SubImageHeight = 1440 ; $AspectRatio = 912/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                8 {$OriginPointElementX = 228 ; $OriginPointElementY = 1440 ;$SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                9 {$OriginPointElementX = 684 ; $OriginPointElementY = 1440 ; $SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}

                            }
                        }
                        4 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1080 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 1800 ; $SubImageWidth = 228 ; $SubImageHeight = 360 ; $AspectRatio = 228/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                7 {$OriginPointElementX = 228 ; $OriginPointElementY = 0 ;$SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                8 {$OriginPointElementX = 684 ; $OriginPointElementY = 0 ; $SubImageWidth = 456 ; $SubImageHeight = 720 ; $AspectRatio = 456/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                9 {$OriginPointElementX = 228 ; $OriginPointElementY = 720 ; $SubImageWidth = 912 ; $SubImageHeight = 1440 ; $AspectRatio = 912/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function Portrait_3_Elements_Base3_Narrow ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 3; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            { @(1,4) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1350 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2700 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitNarrow -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            { @(2,5) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1350 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitNarrow -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2490 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            { @(3,6) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitNarrow -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1140 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2490 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitWide -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Portrait_4_Elements_Base3_Even {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )


    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('PortraitEven')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        if ($SubLayoutType -eq 'PortraitEven') {$SubLayout = Get-Random -Minimum 1 -Maximum 5}

        for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutType)
            {
                'PortraitEven' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 427 ; $OriginPointElementY = 0 ;$SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 853 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 853 ; $OriginPointElementY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 853 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 427 ; $OriginPointElementY = 1440 ;$SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 853 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 853 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 853 ; $OriginPointElementY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ;$SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 427 ; $OriginPointElementY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 427 ; $OriginPointElementY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 853 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        4 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 427 ; $OriginPointElementY = 0 ;$SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 853 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 427 ; $OriginPointElementY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function Portrait_4_Elements_Base3_Even ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 3; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            { @(1..6) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitEven -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitEven -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType PortraitEven -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Mixed_1_6xElements_Base6Thirds_15Px6L {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    
    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$true)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('Small','Large')]
            [string]$SubLayoutSize,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        for ($SubLayoutNum=1; $SubLayoutNum -le 4; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutSize)
            {
                'Small' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 380 ; $OriginPointElementY = 0 ;$SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 760 ; $OriginPointElementY = 0 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 760 ; $OriginPointElementY = 360 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 380 ; $OriginPointElementY = 0 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 380 ; $OriginPointElementY = 360 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 520 ; $OriginPointElementY = 0 ; $SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 380 ; $SubImageHeight = 720 ; $AspectRatio = 380/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 520 ; $SubImageHeight = 360 ; $AspectRatio = 520/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        { @(4,5,6) -contains $_ } {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 427 ; $OriginPointElementY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 853 ; $OriginPointElementY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'Large'
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 760 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1520 ; $OriginPointElementY = 0 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1520 ; $OriginPointElementY = 720 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 760 ; $OriginPointElementY = 0 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1800 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 760 ; $OriginPointElementY = 720 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 1040 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1800 ; $OriginPointElementY = 0 ; $SubImageWidth = 760 ; $SubImageHeight = 1440 ; $AspectRatio = 760/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 1040 ; $SubImageHeight = 720 ; $AspectRatio = 1040/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        { @(4,5,6) -contains $_ } {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 853 ; $OriginPointElementY = 0 ; $SubImageWidth = 854 ; $SubImageHeight = 1440 ; $AspectRatio = 854/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1707 ; $OriginPointElementY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    } 
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function New-Wallpaper_Mixed_1_6xElements_Base6Thirds_15Px6L body ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    $SubLayoutList = @(1..6) # Permutations within function Set-SubLayout

    for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        $SubLayout = Get-Random -InputObject $SubLayoutList

        $SubLayoutList = $SubLayoutList | Where-Object {$_ -ne $SubLayout}
        
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            1 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            2 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            3 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            4 {            
                switch ($SubLayoutNum)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Large -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    6 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -SubLayout $SubLayout -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutSize Small -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Mixed_2_Thirds {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }

    for ($i=1; $i -le 14; $i++) 
    {
        switch ($Layout) # Intended for 3840 pixels wide x 2160 pixels high (standard 4K)
        {
            1 {            
                switch ($i)
                {                    
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 853 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    4 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 1707 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 2133 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    7 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    8 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 853 ; $OriginPointY = 720 ; $SubImageWidth = 854 ; $SubImageHeight = 1440 ; $AspectRatio = 854/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 1707 ; $OriginPointY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    11 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    12 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 2987 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }
            2 {            
                switch ($i)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 1707 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    4 {$OriginPointX = 2133 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 2987 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    7 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    8 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    9 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 2133 ; $OriginPointY = 720 ; $SubImageWidth = 854 ; $SubImageHeight = 1440 ; $AspectRatio = 854/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2987 ; $OriginPointY = 720 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    12 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 427 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 853 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }
            3 {            
                switch ($i)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 427 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 853 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 2133 ; $OriginPointY = 0 ; $SubImageWidth = 854 ; $SubImageHeight = 1440 ; $AspectRatio = 854/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 2987 ; $OriginPointY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    7 {$OriginPointX = 0 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    8 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}
                    
                    9 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 1707 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 2133 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    12 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 2987 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    14 {$OriginPointX = 3413 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    default {break}
                }
            }
            4 {            
                switch ($i)
                {
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    2 {$OriginPointX = 853 ; $OriginPointY = 0 ; $SubImageWidth = 854 ; $SubImageHeight = 1440 ; $AspectRatio = 854/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    3 {$OriginPointX = 1707 ; $OriginPointY = 0 ; $SubImageWidth = 853 ; $SubImageHeight = 1440 ; $AspectRatio = 853/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    4 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    5 {$OriginPointX = 2987 ; $OriginPointY = 0 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    6 {$OriginPointX = 3413 ; $OriginPointY = 0 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}

                    7 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    8 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    9 {$OriginPointX = 427 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    10 {$OriginPointX = 853 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    11 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    12 {$OriginPointX = 1707 ; $OriginPointY = 1440 ; $SubImageWidth = 426 ; $SubImageHeight = 720 ; $AspectRatio = 426/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    13 {$OriginPointX = 2133 ; $OriginPointY = 1440 ; $SubImageWidth = 427 ; $SubImageHeight = 720 ; $AspectRatio = 427/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true}
                    
                    14 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $SubImageWidth = 1280 ; $SubImageHeight = 720 ; $AspectRatio = 1280/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false}

                    default {break}
                }
            }
        }
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
        
        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        # Add SubImageObject to WallpaperObject, after cropping to acheive the AspectRatio and applying RandomFlipHorizontal
        $WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal
    }

    return $WallpaperObject
}


function New-Wallpaper_Mixed_3_Element_Base3 {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )


    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('Portrait','Landscape')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        if ($SubLayoutType -eq 'Portrait') {$SubLayout = Get-Random -Minimum 1 -Maximum 5}
        else {$SubLayout = Get-Random -Minimum 1 -Maximum 4} # Landscape

        for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutType)
            {
                'Portrait' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 900 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 450 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 450 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 900 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        4 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 450 ; $OriginPointElementY = 0 ;$SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 900 ; $OriginPointElementY = 0 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 0 ; $OriginPointElementY = 1440 ; $SubImageWidth = 450 ; $SubImageHeight = 720 ; $AspectRatio = 450/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                6 {$OriginPointElementX = 450 ; $OriginPointElementY = 720 ; $SubImageWidth = 900 ; $SubImageHeight = 1440 ; $AspectRatio = 900/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'Landscape'
                {
                    switch ($SubLayout)
                    {  
                        1 { # Small
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 570 ; $SubImageHeight = 360 ; $AspectRatio = 570/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 570 ; $OriginPointElementY = 0 ; $SubImageWidth = 570 ; $SubImageHeight = 360 ; $AspectRatio = 570/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ; $SubImageWidth = 570 ; $SubImageHeight = 360 ; $AspectRatio = 570/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 570 ; $OriginPointElementY = 360 ; $SubImageWidth = 570 ; $SubImageHeight = 360 ; $AspectRatio = 570/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        { @(2,3) -contains $_ } { # Large
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1140 ; $SubImageHeight = 720 ; $AspectRatio = 1140/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    } 
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function Mixed_3_Element_Base3 ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 5; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            { @(1,4) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1350 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2700 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2700 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2700 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            { @(2,5) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1350 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 1350 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1350 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2490 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            { @(3,6) -contains $_ } {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Landscape -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1140 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2490 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Portrait -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Mixed_4_LargeLandscape {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    
    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('Block1','Block2','Large3','Long4')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal

        if ($SubLayoutType -eq 'Block1') {$SubLayoutTypeSwitch = (Get-Random -InputObject @('Block1a','Block1b'))}
        elseif ($SubLayoutType -eq 'Block2') {$SubLayoutTypeSwitch = $SubLayoutType}
        elseif ($SubLayoutType -eq 'Large3') {$SubLayoutTypeSwitch = (Get-Random -InputObject @('Large3a','Large3b','Large3c','Large3d'))}
        elseif ($SubLayoutType -eq 'Long4') {$SubLayoutTypeSwitch = (Get-Random -InputObject @('Long4a','Long4b'))}
                
        for ($SubLayoutNum=1; $SubLayoutNum -le 6; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutTypeSwitch)
            {
                'Large3a' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 1600 ; $SubImageHeight = 960 ; $AspectRatio = 1600/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 800 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 1600 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        5 {$OriginPointElementX = 1600 ; $OriginPointElementY = 480 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        6 {$OriginPointElementX = 1600 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Large3b' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1600 ; $SubImageHeight = 960 ; $AspectRatio = 1600/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 800 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 1600 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        5 {$OriginPointElementX = 1600 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        6 {$OriginPointElementX = 1600 ; $OriginPointElementY = 480 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Large3c' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 800 ; $OriginPointElementY = 0 ; $SubImageWidth = 1600 ; $SubImageHeight = 960 ; $AspectRatio = 1600/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 800 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 1600 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        5 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        6 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Large3d' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 800 ; $OriginPointElementY = 480 ; $SubImageWidth = 1600 ; $SubImageHeight = 960 ; $AspectRatio = 1600/960 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 800 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 1600 ; $OriginPointElementY = 0 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        5 {$OriginPointElementX = 0 ; $OriginPointElementY = 480 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        6 {$OriginPointElementX = 0 ; $OriginPointElementY = 960 ; $SubImageWidth = 800 ; $SubImageHeight = 480 ; $AspectRatio = 800/480 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Long4a' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 400 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 800 ; $OriginPointElementY = 0 ; $SubImageWidth = 1200 ; $SubImageHeight = 720 ; $AspectRatio = 1200/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 2000 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Long4b' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 400 ; $OriginPointElementY = 0 ; $SubImageWidth = 1200 ; $SubImageHeight = 720 ; $AspectRatio = 1200/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 1600 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 2000 ; $OriginPointElementY = 0 ; $SubImageWidth = 400 ; $SubImageHeight = 720 ; $AspectRatio = 400/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Block1a' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 630 ; $SubImageHeight = 1080 ; $AspectRatio = 630/1080 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 630 ; $OriginPointElementY = 0 ; $SubImageWidth = 810 ; $SubImageHeight = 540 ; $AspectRatio = 810/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 630 ; $OriginPointElementY = 540 ; $SubImageWidth = 810 ; $SubImageHeight = 540 ; $AspectRatio = 810/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Block1b' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 810 ; $SubImageHeight = 540 ; $AspectRatio = 810/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 0 ; $OriginPointElementY = 540 ; $SubImageWidth = 810 ; $SubImageHeight = 540 ; $AspectRatio = 810/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $false ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 810 ; $OriginPointElementY = 0 ; $SubImageWidth = 630 ; $SubImageHeight = 1080 ; $AspectRatio = 630/1080 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
                'Block2' 
                {
                    switch ($SubLayoutNum)
                    {
                        1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 360 ; $SubImageHeight = 540 ; $AspectRatio = 360/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        2 {$OriginPointElementX = 0 ; $OriginPointElementY = 540 ; $SubImageWidth = 360 ; $SubImageHeight = 540 ; $AspectRatio = 360/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        3 {$OriginPointElementX = 360 ; $OriginPointElementY = 0 ; $SubImageWidth = 720 ; $SubImageHeight = 1080 ; $AspectRatio = 720/1080 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        4 {$OriginPointElementX = 1080 ; $OriginPointElementY = 0 ; $SubImageWidth = 360 ; $SubImageHeight = 540 ; $AspectRatio = 360/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        5 {$OriginPointElementX = 1080 ; $OriginPointElementY = 540 ; $SubImageWidth = 360 ; $SubImageHeight = 540 ; $AspectRatio = 360/540 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                        default {break}
                    }
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function New-Wallpaper_Mixed_4_LargeLandscape body ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 4; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            1 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Long4 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Large3 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2400 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block1 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2400 ; $OriginPointY = 1080 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block2 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            2 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Large3 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Long4 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2400 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block2 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 2400 ; $OriginPointY = 1080 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block1 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            3 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block2 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 1080 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block1 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 1440 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Large3 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1440 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Long4 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            4 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block1 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 0 ; $OriginPointY = 1080 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Block2 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 1440 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Long4 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1440 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType Large3 -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
    
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


function New-Wallpaper_Magazine {
    
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataHashtable,
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [System.Drawing.Bitmap]$WallpaperObject
    )

    
    function Set-SubLayout {
 
        Param(
            [Parameter(Mandatory=$true)]
            [hashtable]$DataHashtable,
            [Parameter(Mandatory=$true)]
            [string]$GroupName,
            [Parameter(Mandatory=$false)]
            [Int32]$SubLayout,     
            [Parameter(Mandatory=$true)]
            [ValidateSet('LargeBlock','MediumBlock','SmallBlock')]
            [string]$SubLayoutType,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointX,
            [Parameter(Mandatory=$true)]
            [Int32]$OriginPointY,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperWidth,
            [Parameter(Mandatory=$true)]
            [Int32]$WallpaperHeight,
            [Parameter(Mandatory=$true)]
            [System.Drawing.Image]$WallpaperObject
        )

        $ZoomPercent = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.ZoomPercent
        $RandomFlipHorizontal = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.RandomFlipHorizontal
        
        if ($SubLayoutType -eq 'LargeBlock') {$SubLayout = Get-Random -Minimum 1 -Maximum 4}
        elseif ($SubLayoutType -eq 'MediumBlock') {$SubLayout = Get-Random -Minimum 1 -Maximum 3}
        else {$SubLayout = Get-Random -Minimum 1 -Maximum 4} # SmallBlock
        
        for ($SubLayoutNum=1; $SubLayoutNum -le 5; $SubLayoutNum++) # SubLayoutNum at a higher hierarchial level to reduce for loops.
        {
            switch ($SubLayoutType)
            {
                'LargeBlock' 
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 1024 ; $OriginPointElementY = 0 ;$SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 2048 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 2048 ; $OriginPointElementY = 720 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                             {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 1024 ; $OriginPointElementY = 0 ;$SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1024 ; $OriginPointElementY = 720 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1536 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                              {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ;$SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 512 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1536 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'MediumBlock'
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ;$SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 0 ; $OriginPointElementY = 720 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 0 ; $OriginPointElementY = 1080 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 256 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                             {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 1024 ; $SubImageHeight = 1440 ; $AspectRatio = 1024/1440 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 1024 ; $OriginPointElementY = 0 ;$SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1024 ; $OriginPointElementY = 360 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1024 ; $OriginPointElementY = 720 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                5 {$OriginPointElementX = 1024 ; $OriginPointElementY = 1080 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    }
                }
                'SmallBlock'
                {
                    switch ($SubLayout)
                    {  
                        1 {
                            switch ($SubLayoutNum)
                            {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 0 ; $OriginPointElementY = 360 ;$SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 256 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 768 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        2 {
                            switch ($SubLayoutNum)
                             {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 512 ; $OriginPointElementY = 0 ;$SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 512 ; $OriginPointElementY = 360 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 768 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                        3 {
                            switch ($SubLayoutNum)
                              {
                                1 {$OriginPointElementX = 0 ; $OriginPointElementY = 0 ; $SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                2 {$OriginPointElementX = 512 ; $OriginPointElementY = 0 ;$SubImageWidth = 512 ; $SubImageHeight = 720 ; $AspectRatio = 512/720 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                3 {$OriginPointElementX = 1024 ; $OriginPointElementY = 0 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                4 {$OriginPointElementX = 1024 ; $OriginPointElementY = 360 ; $SubImageWidth = 256 ; $SubImageHeight = 360 ; $AspectRatio = 256/360 ; $SubImageObject = Get-SubImageObject -DataHashtable $DataHashtable -GroupName $GroupName -SelectPortrait $true ; if ($SubImageObject) {$WallpaperObject = Set-SubImage -SubImageObject $SubImageObject -WallpaperObject $WallpaperObject -OriginPointX ($OriginPointX + $OriginPointElementX) -OriginPointY ($OriginPointY + $OriginPointElementY) -SubImageWidth $SubImageWidth -SubImageHeight $SubImageHeight -AspectRatio $AspectRatio -ZoomPercent $ZoomPercent -RandomFlipHorizontal $RandomFlipHorizontal}}
                                default {break}
                            }
                        }
                    } 
                }
            }
        }

        return $WallpaperObject
    }

    
    ####################  function New-Wallpaper_Magazine body ####################

    # Derive $WallpaperLayout from the function name
    [string]$WallpaperLayout = "$($MyInvocation.MyCommand.Name)".Replace('New-Wallpaper_','')

    if ($DataHashtable.PictureNameOverride -gt 0)
    {
        $Layout = Get-Random -Minimum 1 -Maximum ($DataHashtable.WallpaperLayoutCountHash.$WallpaperLayout + 1)
    }
    else
    {
        $Layout = $DataHashtable.LayoutTracker
    }
    
    for ($SubLayoutNum=1; $SubLayoutNum -le 5; $SubLayoutNum++) # SubLayout at a higher hierarchial level to reduce the number of for loops
    {
        switch ($Layout) # Wallpaper image Layout from the main script body
        {
            1 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType LargeBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType MediumBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            2 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType MediumBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType LargeBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            3 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType MediumBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 2560 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 0 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 1280 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType LargeBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
            4 {            
                switch ($SubLayoutNum)
                { 
                    1 {$OriginPointX = 0 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    2 {$OriginPointX = 1280 ; $OriginPointY = 0 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType LargeBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    3 {$OriginPointX = 0 ; $OriginPointY = 720 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType MediumBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    4 {$OriginPointX = 1280 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    5 {$OriginPointX = 2560 ; $OriginPointY = 1440 ; $WallpaperObject = Set-SubLayout -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperWidth $WallpaperWidth -WallpaperHeight $WallpaperHeight -SubLayoutType SmallBlock -OriginPointX $OriginPointX -OriginPointY $OriginPointY -WallpaperObject $WallpaperObject}
                    default {break}
                }
            }
        } 
        
        # Return early if one of the Get-SubImageObject calls fail
        if ($DataHashtable.GetSubImageObjectSuccess -eq $false) {return $WallpaperObject}
    }

    return $WallpaperObject
}


########################################################################## Script Body ##########################################################################

# Using .Net Picture Drawing Assembly
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$DataHashtable = @{}
$DataHashtable = Get-InputJsonFileData -InputsJsonFile $InputsJsonFile -DataHashtable $DataHashtable -WallpaperDimensions $WallpaperDimensions -WallpaperLayoutCountHash $WallpaperLayoutCountHash -Thread $Thread -AutoMultiThreadActive $AutoMultiThreadActive  # Pass InputsJsonFile and WallpaperLayoutCountHash and optional script parameters during bootstrap

if (-not (Test-Path -Path $DataHashtable.OutputFolder))
{
    do
    {
        $DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable

        sleep -Seconds 5
    }
    until (Test-Path -Path $DataHashtable.OutputFolder)
}

# Infinite loop of Wallpaper image creation
:InfiniteWallpaperLoop do
{
    # Delete any high-numbered Wallpaper files which are not used in the current runtime
    if (-not $Thread)
    {
        if ((Get-ChildItem -Path $DataHashtable.OutputFolder).Name)
        {
            $HighestNumberWallpaperFileName = [Int32]((Get-ChildItem -Path $DataHashtable.OutputFolder).Name | ForEach-Object {$_.Replace('.png','').Replace('_temp','')} | measure -Maximum).Maximum
        }

        if ($DataHashtable.MaxLayouts -lt $HighestNumberWallpaperFileName)
        {
            (($DataHashtable.MaxLayouts + 1)..$HighestNumberWallpaperFileName) | ForEach-Object {if ([System.IO.File]::Exists("$($DataHashtable.OutputFolder)\$_.PNG")) {Remove-Item -Path "$($DataHashtable.OutputFolder)\$_.PNG"}}
        }
    }
    
    # Thread is only set within ThreadJobs
    if ($Thread)
    {
        $DataHashtable.Thread = $Thread  # Used for SleepTime
    }
    elseif ($DataHashtable.NumberThreads -gt 1)
    {
        $ThreadJobsStarted = Start-ThreadThreadJobs -DataHashtable $DataHashtable  # This will fail softly and resume as single thread if PowerShell 7 is not installed

        if ($ThreadJobsStarted)
        {
            $DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable  # Capture $DataHashtable from Get-InputJsonFileData return, $DataHashtable and its pointer reference is re-initialized in-function

            continue InfiniteWallpaperLoop
        }
    }

    ##### Use Default WallpaperLayout, non-Default WallpaperLayoutWeightedList GroupNames are either offline and/or functionally disabled
    if (-not ($DataHashtable.WallpaperLayoutWeightedListEnableState -and $DataHashtable.WeightedWallpaperSourceFoldersOnlineState) -and ($DataHashtable.DefaultFolderEnableState -and $DataHashtable.DefaultSourceFoldersOnlineState))
    {
        # If Thread, Wallpaper names will always be based on Thread
        if ($Thread)
        {
            $DataHashtable.PictureNameOverride = $Thread  # Used for WallpaperLayout functions
        }
        elseif ($DataHashtable.LayoutTracker -ge 1)  # Continue previous LayoutTracker sequence
        {
            $DataHashtable.PictureNameOverride = 0
            $DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts
        }
        else
        # Start new LayoutTracker sequence
        # LayoutTracker is only used for the Default, and will increment through the Default WallpaperLayout variations numberically rather than at random
        {
            $DataHashtable.LayoutTracker = 1
            $DataHashtable.PictureNameOverride = 0
            $DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts
        }
        
        $GroupName = 'Default'
        $WallpaperLayout = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.WallpaperLayout
        $WallpaperFunction = "New-Wallpaper_$WallpaperLayout"
        
        do
        {
            # Default SourceFolder is also offline
            if (-not ($DataHashtable.DefaultFolderEnableState -or $DataHashtable.DefaultSourceFoldersOnlineState))
            {
                Get-WallpaperLayoutWeightedState -DataHashtable $DataHashtable

                sleep 10
            }

            if ((-not $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash) -or ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count -lt 10))
            {
                Get-SourceFileList -DataHashtable $DataHashtable -GroupName $GroupName
            }
            
            # Create Wallpaper image, test for suplus Portrait and Landscape images, and subtract Stopwatch from TimeBetweenCreation
            $Stopwatch = New-WallpaperImage -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperFunction $WallpaperFunction

            # Check for updates to InputsFileJson.json, as well as AutoMultiThread Wallpaper creation exceeding TimeBetweenCreation, using the New-WallpaperImage StopWatch
            if (Test-UpdateRequested -DataHashtable $DataHashtable -Stopwatch $Stopwatch) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
            if ((Test-MultiThread -DataHashtable $DataHashtable -Stopwatch $Stopwatch)) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}

            # Consume excess Portrait or Landscape images if Portrait_Landscape_SurplusBurnoff is enabled in InputsJsonFile.json for the GroupName
            $SurplusWallpaperCreated,$Stopwatch = New-SurplusWallpaperImage -DataHashtable $DataHashtable -GroupName $GroupName -Stopwatch $Stopwatch
            if ($SurplusWallpaperCreated)
            {     
                # Check for updates to InputsFileJson.json, as well as AutoMultiThread Wallpaper creation exceeding TimeBetweenCreation, using the New-WallpaperImage StopWatch
                if (Test-UpdateRequested -DataHashtable $DataHashtable -Stopwatch $Stopwatch) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
                if ((Test-MultiThread -DataHashtable $DataHashtable -Stopwatch $Stopwatch)) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
            }

            # Check the state of SourceFolderOfflineList with Test-SourceFolderState, which will rebuild SourcePictureList if that status changes
            if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList.Count -gt 0)
            {
                $SourceFolderOfflineStateChange,$Void = Test-SourceFolderState -DataHashtable $DataHashtable -GroupName $GroupName

                if ($SourceFolderOfflineStateChange)
                {
                    $DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable
                    
                    Get-SourceFileList -DataHashtable $DataHashtable -GroupName $GroupName

                    $DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts
                }
            }
        
            # If $WallpaperLayoutWeightedListEnableState is enabled, check on $WeightedWallpaperSourceFoldersOnlineState, and move to $WallpaperLayoutWeightedList if 1+ Weighted SourceFolders comes online
            if ($DataHashtable.WallpaperLayoutWeightedListEnableState -and (-not $DataHashtable.WeightedWallpaperSourceFoldersOnlineState))
            {
                Get-WallpaperLayoutWeightedState -DataHashtable $DataHashtable
            }
        }
        While (-not ($DataHashtable.WallpaperLayoutWeightedListEnableState -and $DataHashtable.WeightedWallpaperSourceFoldersOnlineState) -and ($DataHashtable.DefaultFolderEnableState -and $DataHashtable.DefaultSourceFoldersOnlineState))
        # continue to InfiniteWallpaperLoop
    }
    ##### Randomized selection of Weighted WallpaperLayouts, $WallpaperLayoutWeightedListEnableState = $true and $WeightedWallpaperSourceFoldersOnlineState = $true
    elseif ($DataHashtable.WallpaperLayoutWeightedListEnableState -and $DataHashtable.WeightedWallpaperSourceFoldersOnlineState)
    {
        # If Thread, Wallpaper names will always be based on Thread
        if ($Thread)
        {
            $DataHashtable.PictureNameOverride = $Thread  # Used for WallpaperLayout functions
        }
        elseif ($DataHashtable.PictureNameOverride -ge 1)  # Continue previous PictureNameOverride sequence
        {
            $DataHashtable.LayoutTracker = 0
            $DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts
        }
        else  # Start new PictureNameOverride sequence
        {
            $DataHashtable.PictureNameOverride = 1
            $DataHashtable.LayoutTracker = 0
            $DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts
        }

        do
        {
            do
            {
                $i = Get-Random -Minimum 0 -Maximum $DataHashtable.WallpaperLayoutWeightedRuntimeList.Count

                $GroupName = $DataHashtable.WallpaperLayoutWeightedRuntimeList[$i]
                $DataHashtable.WallpaperLayoutWeightedRuntimeList.RemoveAt($i)

                $SelectedLayout = $DataHashtable.WallpaperLayoutWeightedList.$GroupName.WallpaperLayout
                $WallpaperFunction = "New-Wallpaper_$SelectedLayout"

                if ((-not $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash) -or ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourcePictureHash.SourcePictureList.Count -lt 10))
                {
                    Get-SourceFileList -DataHashtable $DataHashtable -GroupName $GroupName
                }
                
                # Create Wallpaper image, test for suplus Portrait and Landscape images, and subtract Stopwatch from TimeBetweenCreation
                $Stopwatch = New-WallpaperImage -DataHashtable $DataHashtable -GroupName $GroupName -WallpaperFunction $WallpaperFunction

                # Check for updates to InputsFileJson.json, as well as AutoMultiThread Wallpaper creation exceeding TimeBetweenCreation, using the New-WallpaperImage StopWatch
                if (Test-UpdateRequested -DataHashtable $DataHashtable -Stopwatch $Stopwatch) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
                if ((Test-MultiThread -DataHashtable $DataHashtable -Stopwatch $Stopwatch)) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}

                # Consume excess Portrait or Landscape images if Portrait_Landscape_SurplusBurnoff is enabled in InputsJsonFile.json for the GroupName
                $SurplusWallpaperCreated,$Stopwatch = New-SurplusWallpaperImage -DataHashtable $DataHashtable -GroupName $GroupName -Stopwatch $Stopwatch

                if ($SurplusWallpaperCreated)
                {     
                    # Check for updates to InputsFileJson.json, as well as AutoMultiThread Wallpaper creation exceeding TimeBetweenCreation, using the New-WallpaperImage StopWatch
                    if (Test-UpdateRequested -DataHashtable $DataHashtable -Stopwatch $Stopwatch) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
                    if ((Test-MultiThread -DataHashtable $DataHashtable -Stopwatch $Stopwatch)) {$DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable ; continue InfiniteWallpaperLoop}
                }

                ## Check the state of SourceFolderOfflineList with Test-SourceFolderState, which will rebuild SourcePictureList if that status changes
                if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList.Count -gt 0)
                {
                    $SourceFolderOfflineStateChange,$Void = Test-SourceFolderState -DataHashtable $DataHashtable -GroupName $GroupName
                }
        
                # Ensure a $SourceFolderOfflineStateChange has not resulted in WeightedWallpaperSourceFoldersOnlineState, otherwise rebuild the GroupName SourcePictureHash
                if ($SourceFolderOfflineStateChange)
                {
                    if (-not $Thread) {$DataHashtable.OutputFolderFastRefresh = $DataHashtable.MaxLayouts}
                    
                    $DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable
                    
                    Get-SourceFileList -DataHashtable $DataHashtable -GroupName $GroupName
                }
                
                $SourceFolderOfflineStateChange = $false  # Reset for next loop

                # Compare offline SourceFolders with SourceFolder count
                if ($DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolderOfflineList.Count -ge $DataHashtable.WallpaperLayoutWeightedList.$GroupName.SourceFolders.Count)
                {
                    $DataHashtable.WallpaperLayoutWeightedRuntimeList = [System.Collections.Generic.List[string]]($DataHashtable.WallpaperLayoutWeightedRuntimeList | Where-Object {$_ -ne $GroupName})

                    Get-WallpaperLayoutWeightedState -DataHashtable $DataHashtable
                }
            }
            While (($DataHashtable.WallpaperLayoutWeightedRuntimeList.Count -gt 0) -and $DataHashtable.WallpaperLayoutWeightedListEnableState -and $DataHashtable.WeightedWallpaperSourceFoldersOnlineState)

            Set-WallpaperWeightedLayoutList -DataHashtable $DataHashtable
        } 
        While ($DataHashtable.WallpaperLayoutWeightedListEnableState -and $DataHashtable.WeightedWallpaperSourceFoldersOnlineState)
        # continue to InfiniteWallpaperLoop
    }
    else
    {
        $DataHashtable = Get-InputJsonFileData -DataHashtable $DataHashtable

        sleep -Seconds 5
    }
    # continue to InfiniteWallpaperLoop
}
While ($true)