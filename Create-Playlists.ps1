#Creates a UTF8 encoded m3u8 playlist for media files in provided folder path 
param (
	[string]$d,
	[String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv', '.mpg', '.ts')
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

function Create-Playlists-For-Folders($folderPath){
	$dirs = @(Get-ChildItem -literalPath $folderPath | Where-Object{ $_.PSIsContainer })
	
	foreach($d in $dirs){
		try {
			&util.Create-Playlist-m3u -d $d.FullName -t $t	
			write-host ("CREATED M3U FOR: " + $d.Name)
		}
		catch {
			write-host "ERROR occurred for: " + $d.FullName
  			Write-Host $_.ScriptStackTrace
		}
		
	}
}

#Check if help was invoked
if($help.ispresent){Get-Help-Info}

#Check if source dir was provided -> ask for source
while(-not $d){$d = $(Read-Host 'Source Folder')}
if(!(Test-Path -literalPath $d -PathType Container)) {throw $NonExisting_Folder_Error}

Create-Playlists-For-Folders $d 
