#Creates a UTF8 encoded m3u8 playlist for media files in provided folder path 
param (
	[string]$d,
	[String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv', '.mpg', '.ts'),
	[switch]$help,
	[switch]$recurse,
	[switch]$overwrite
 )
 
$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"
$NonExisting_Folder_Error = New-Object System.FormatException "-d provided folder does not exist!"

function Get-Help-Info{
    write-host("`nSYNTAX")
	write-host("  " + "Create-Playlist-m3u8.ps1 [[-d]<string>] [[-t]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v')>] [[-recurse]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Provides the dir path to be searched for media files to convert.")
	write-host("  " + "{0,-15} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-recurse", "Search directory recursively for files.")
	write-host("  " + "{0,-15} {1}" -f "-overwrite", "Overwrites existing playlists")
	exit
}

filter Get-Files-Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}
	
function Get-Script-Folder-Path{
  return Split-Path -Path $script:MyInvocation.MyCommand.Path -Parent
}

function Create-Playlist($folderPath, $r){
	try {
		$files = @(Get-ChildItem -recurse -literalPath $folderPath | Get-Files-Where-Extension $t )	
	}
	catch {
		write-host "ERROR occurred for:" + $folderPath
  		Write-Host $_.ScriptStackTrace
		exit
	}
	
	$folder = Get-Item $folderPath
	$playlistPath = Join-Path $folder.FullName ($folder.BaseName + ".m3u")
	$playlistFile = New-Item -Path $playlistPath -ItemType File -Force
	
	#"#EXTM3U" | Out-File -Encoding UTF8 -Append -FilePath $playlistPath
	#"`n" | Out-File -Encoding UTF8 -Append -FilePath $playlistPath
	foreach($file in $files){
		#("#EXTINF 0," + $file.basename + " - " + $folder.BaseName) | Out-File -Encoding UTF8 -Append -FilePath $playlistPath
		("." + $file.FullName.Replace($folder.FullName, "")) | Out-File -Encoding UTF8 -Append -FilePath $playlistPath
	}

	if($r){
		$dirs = @(Get-ChildItem -literalPath $folderPath | Where-Object{ $_.PSIsContainer })
		if($dirs.Count -gt 0) {
			foreach($d in $dirs){
				$mediaFileCount = @(Get-ChildItem -literalPath $d.FullName | Get-Files-Where-Extension $t )
				if($mediaFileCount.Count -gt 0){
					Create-Playlist $d.FullName
				}
			}
		}
	}
}

#Check if help was invoked
if($help.ispresent){Get-Help-Info}

#Check if source dir was provided -> ask for source
while(-not $d){$d = $(Read-Host 'Source Folder')}
if(!(Test-Path -literalPath $d -PathType Container)) {throw $NonExisting_Folder_Error}

Create-Playlist $d $recurse
