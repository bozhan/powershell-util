# Adding Utils to Profile
By adding the util scripts to the powershell profile, one can access all util scripts globally by just using util.<script name>.  
The powershell profile file can be found in C:\Users\%USERNAME%\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1  
You can define a repo variable and add all util scripst you plan on using.  
```
$ps_repo_dir = "C:\Users\BoZ\dev\powershell"
New-Alias util.Get-Media-Info $ps_repo_dir\util\Get-Media-Info.ps1
New-Alias util.Extract-Files-From-Folders $ps_repo_dir\util\Extract-Files-From-Folders.ps1
...
```
or in case you want your profile to automatically load all util scripst on startup, you can do the following  
```
$ps_util_path = "<PATH TO REPOSITORY>"
$utils = @(Get-ChildItem -literalPath $ps_util_path | where {$_.extension -eq '.ps1'})
foreach($util in $utils){
	$utilAlias = "util." + $util.basename
	New-Alias -Name $utilAlias -Value $util.FullName
}
Clear-Variable -Name "utilAlias" -Scope Global
Clear-Variable -Name "util" -Scope Global
Clear-Variable -Name "utils" -Scope Global
```

# powershell-util
Various powershell utility functions and scripts used in other poweshell project.

# Get-MediaInfo
Reports a summary of the video files in each folder found in the provided path.
Recurse switch -r could be used here if you want to report on subfolders.

The summary consists of several attributes:
MB/Min - ratio of size in MB to lengt of the video (helful if you want to figure out if a media file can be compressed further with e.g. h264/h265 without significant quality loss)
Size(MB) - size of the folder containing media files (only media files are considered in this calcualtion) 
Length(Min) - length of the video in minutes
#Files - number of files in each reported folder
Video BR - average video bitrate for all files in a listed folder
Audio BR - average audio bitrate for all files in a listed folder
Frame Width - the average video fram width for all files in a listed folder
Forlder Name - the folder name listed
Forlder Parent - the parent folder of the listed folder name

For more information on the syntax and options avaialbe use the -help switch on the cmdlet.
