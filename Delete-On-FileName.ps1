 param (
    [string]$d,
    [String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv', '.ts', '.webm'),
    [string]$s,
	[switch]$right,
	[switch]$left,
	[switch]$mid,
	[switch]$help,
	[switch]$exclude,
	[switch]$report,
	[switch]$r
 )
 
$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"
$Missing_SearchString_Error = New-Object System.FormatException "-s (search string) is missing!"
$Missing_SearchStringPosition_Error = New-Object System.FormatException "You have to specify at least one switch for position of searched string (-right, -left, -mid)"
 
filter Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Delete-On-FileName [[-d]<string>] [[-t]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v')>] [[-s]<string>] [[-r]<switch>] [[-right]<switch>] [[-left]<switch>] [[-mid]<switch>] [[-exclude]<switch>] [[-report]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-20} {1}" -f "-d", "Specify directrory to containing media files.")
	write-host("  " + "{0,-20} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-20} {1}" -f "-s", "Set string contained into the file names searched for.")
	write-host("  " + "{0,-20} {1}" -f "-right/-mid/-left", "At least one switch for position of searched string has to be given.")
	write-host("  " + "{0,-20} {1}" -f "-exclude", "Set to exclude the matched filenames from the search/delete.")
	write-host("  " + "{0,-20} {1}" -f "-r", "Search directory recursively for files.")
	write-host("  " + "{0,-20} {1}" -f "-reportOnly", "Will only report which files would be deleted.")
	exit
}

while(-not $d){$d = $(Read-Host 'Source Folder')}
if(!(Test-Path -literalPath $d -PathType Container)) {throw $Missing_Folder_Error}
if (-not ($right.ispresent -or $left.ispresent -or $mid.ispresent)){$mid=$true} #throw $Missing_SearchStringPosition_Error}
if (-not $s){$s = "*" }#throw $Missing_SearchString_Error}

if ($r.IsPresent){
	$files = @(Get-ChildItem -literalPath $d -recurse | Where-Extension $t)
}else{
	$files = @(Get-ChildItem -literalPath $d | Where-Extension $t)
}

if ($right.ispresent) { 
	$searchStr = '*' + $s
}elseif ($left.ispresent){
	$searchStr = $s + '*'
}elseif ($mid.ispresent){
	$searchStr = '*' + $s + '*'
}

if ($exclude.isPresent) {
	$filterFiles = @($files | ? {!($_.BaseName -like $searchStr)})
}else{
	$filterFiles = @($files | ? {($_.BaseName -like $searchStr)})
}

Write-Host "Source:" $d
Write-Host "File count:" $filterFiles.count

Add-Type -AssemblyName Microsoft.VisualBasic

foreach ($file in $filterFiles){
	$i = $i + 1	
	if ($report.isPresent){
		write-host $i "/" $filterFiles.count " MARKED " $file.Name
	}else{
		write-host $i "/" $filterFiles.count " DELETED " $file.Name
		[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($file.FullName,'OnlyErrorDialogs','SendToRecycleBin')
	}
}





    

