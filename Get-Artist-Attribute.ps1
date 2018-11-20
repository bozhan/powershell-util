param (
	[string]$d,
	[String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.mkv', '.mp3'),
	[string]$maxBitrate = 0,
	[switch]$help,
	[switch]$csv = $false,
	[switch]$r
)

$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"
 
filter Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

function Get-AttributeValue($file, [string]$attrName){
	$shellObject = New-Object -ComObject Shell.Application
	$directoryObject = $shellObject.NameSpace($file.Directory.FullName)
	$fileObject = $directoryObject.ParseName($file.Name)
			
	# Find the index attribute with give name.
	for($index = 1; $index -lt 300; ++$index) {
		$name = $directoryObject.GetDetailsOf($directoryObject.Items, $index)
		if($name -eq $attrName) {
			$attrIndex = $index 
			break
		}
	}
			
	# Get the attr of the file.
	$attrString = $directoryObject.GetDetailsOf($fileObject, $attrIndex)
	if( $attrString -match '\d+' ) { 
		[int]$attrValue = $matches[0] 
	}else { 
		[int]$attrValue = -1 
	}
	
	return $attrValue
}

function Get-ArtistAttributesForFiles([string]$folderPath, [boolean]$recurse = $r){
	if($recurse){
		$files = @(Get-ChildItem -literalPath $folderPath -recurse | Where-Extension $t)
	}else{
		$files = @(Get-ChildItem -literalPath $folderPath | Where-Extension $t)
	}
	
	if($files.count -gt 0){
		foreach($file in $files){
			$artist = Get-AttributeValue $file "Authors" 
			write-host("{0,-5} {1}" -f $artist, $file.Name)
		}
	}else{
		write-host($files.count + " - found in " + $folderPath)
	}
}


if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Get-MeanBitrate [[-d]<string>] [[-t]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v')>] [[-maxBitrate]<integer=0>] [[-r]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Specify directrory to containing media files.")
	write-host("  " + "{0,-15} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-r", "Search directory recursively for files.")
	exit
}

$d = $(Read-Host 'Source Folder')
if (-not $d){throw $Missing_Folder_Error}

Get-ArtistAttributesForFiles $d $r