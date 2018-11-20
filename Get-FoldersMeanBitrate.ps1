 #[string]::join("`t", (0..3,15,17 | % {$fields[$_]}))
 
 #composite formatting
 #https://docs.microsoft.com/en-us/dotnet/standard/base-types/composite-formatting
 #<string> -f { index[,alignment][:formatString]} 
 
 #table data structure
 #https://www.petri.com/dancing-on-the-table-with-powershell
 
 param (
    [string]$d,
    [String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv'),
    [string]$maxBitrate = 0,
		[switch]$help,
		[switch]$r
 )

$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"
$totalBrName = "Total bitrate"
$totalABrName = "Bit rate"
$totalFwName = "Frame width"
 
filter Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

function Get-MeanAttribute([string]$folderPath, [string]$attrName){
	$shellObject = New-Object -ComObject Shell.Application
	$files = @(Get-ChildItem -literalPath $folderPath -recurse | Where-Extension $t)
	$total = 0
	
	if($files.count -gt 0){
		foreach($file in $files){
			$directoryObject = $shellObject.NameSpace( $file.Directory.FullName )
			$fileObject = $directoryObject.ParseName( $file.Name )
			
			# Find the index attribute with give name.
			for( $index = 1; $index  -lt 300; ++$index ) {
				$name = $directoryObject.GetDetailsOf( $directoryObject.Items, $index )
				if($name -eq $attrName) {
					$attrIndex = $index 
					break
				}
			}
			
			# Get the attr of the file.
			$attrString = $directoryObject.GetDetailsOf( $fileObject, $attrIndex )
			if( $attrString -match '\d+' ) { 
				[int]$attrValue = $matches[0] 
			}else { 
				$attrValue = -1 
			}
			
			$total += $attrValue
		}
		$result = $total/$files.count
	}else{
		$result = 0
	}
	return $result
}

function Get-MeanBitrateDataOfFolders([string]$folderPath, [int]$mbr = "$maxBitrate", [boolean]$recurse = $r, [System.Data.DataTable]$table){
	$folders = @(Get-ChildItem -literalPath $folderPath | ?{ $_.PSIsContainer })
	$progress = 0
	
	if($folders.count -gt 0){
		foreach($folder in $folders) {		
			Write-Progress -Activity "Search folders" -Status "$progress% Complete:" -PercentComplete $progress
			if ($recurse){
				Get-MeanBitrateDataOfFolders $folder.Fullname $mbr $false ([ref]$table)
			}	
			$vbr = Get-MeanAttribute $folder.Fullname $totalBrName
			if($vbr -ge $mbr) {
				$row = $table.NewRow()
				$size = Get-ChildItem -literalpath $folder.fullname -r | Measure-Object -property length -sum
				$row["size"] = [math]::Round($size.sum / "1MB")
				$row["vbr"] = $vbr
				$row["abr"] = Get-MeanAttribute $folder.Fullname $totalABrName 
				$row["width"] = Get-MeanAttribute $folder.Fullname $totalFwName 
				$row["path"] = $folder.Name
				$row["parent"] = (get-item $folder.Fullname ).parent.Name
				$table.Rows.Add($row)
			}
			$i = $i +1
			$progress = [math]::Round(($i/$folders.count)*100)
		}	
	}
}

function Get-MeanBitrateOfFolders([string]$folderPath, [int]$mbr = "$maxBitrate", [boolean]$recurse = $r){
	$table = New-Object System.Data.DataTable
	$table.Columns.Add("size","string") | Out-Null
	$table.Columns.Add("vbr","int32") | Out-Null
	$table.Columns.Add("abr","int32") | Out-Null
	$table.Columns.Add("width","int32") | Out-Null
	$table.Columns.add("path","string") | Out-Null
	$table.Columns.add("parent","string") | Out-Null

	Get-MeanBitrateDataOfFolders $folderPath $mbr $recurse ([ref]$table)
	
	$properties = @(
    @{Name=" Size(MB) ";Expression={$_["size"]};Alignment='Center'}
		@{Name=" Video BR ";Expression={$_["vbr"]};Alignment='Center'}
		@{Name=" Audio BR ";Expression={$_["abr"]};Alignment='Center'}
		@{Name=" Frame Width ";Expression={$_["width"]};Alignment='Center'}
		@{Name=" Forlder Name ";Expression={$_["path"]};Alignment='Left'}
		@{Name=" Forlder Parent ";Expression={$_["parent"]};Alignment='Left'}
	)
	$table | Format-Table -Property $properties -AutoSize -Wrap
}

if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Get-MeanBitrate [[-d]<string>] [[-t]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v')>] [[-maxBitrate]<integer=0>] [[-r]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Specify directrory to containing media files.")
	write-host("  " + "{0,-15} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-maxBitrate", "Set maximum bitrate, all avg bitrate larger than the set value will be reported.")
	write-host("  " + "{0,-15} {1}" -f "-r", "Search directory recursively for files.")
	exit
}

$d = $(Read-Host 'Source Folder')
if (-not $d){throw $Missing_Folder_Error}

if ($r){
	$detail = $true
}else{
	$detail = $false
}

Get-MeanBitrateOfFolders $d $maxBitrate $detail