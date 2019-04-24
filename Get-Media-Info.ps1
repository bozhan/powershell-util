#Additonal references
#[string]::join("`t", (0..3,15,17 | % {$fields[$_]}))

#composite formatting
#https://docs.microsoft.com/en-us/dotnet/standard/base-types/composite-formatting
#<string> -f { index[,alignment][:formatString]} 

#table data structure
#https://www.petri.com/dancing-on-the-table-with-powershell

#working with dates and time
#$span = New-TimeSpan -Hours $hours -Minutes $min -Seconds $sec
#$span = New-TimeSpan -Hours $attrValue.Hour -Minutes $attrValue.Minutes -Seconds $attrValue.Seconds
#$span.Hour = $attrValue.Hour
#$span.Minutes = $attrValue.Minute
#$span.Seconds = $attrValue.Second
#write-host $span
 
 param (
    [string]$d,
    [String[]]$t = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv'),
    [string]$sort,
	[switch]$wrap, 
	[switch]$help,
	[switch]$r,
	[String[]]$filter,
	[switch]$tofile
 )

 function Get-Help-Message{
	write-host("`nSYNTAX")
	write-host("  " + 
		"Get-MeanBitrate [[-d]<string>] [[-t]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v')>]`n" + 
		"[[-r]<switch>] [-sort<string[]=(filter/sort attributes)>]`n" + 
		"[[-filter]<string>] [<string[]=('filter_sort_attribute1 > min_value1', 'table_attribute2 = min_value2', ...)>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Specify directrory to containing media files.")
	write-host("  " + "{0,-15} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-maxBitrate", "Set maximum bitrate, all avg bitrate larger than the set value will be reported.")
	write-host("  " + "{0,-15} {1}" -f "-r", "Search directory recursively for files.")
	write-host("  " + "{0,-15} {1}" -f "-filter", "applies additive filter to output. The filter is defined in a string tuple as shown in the syntax.`n" +
		"use one of the compareson operators between attribute and values provided (>, <, =, <=, >=")
	write-host("  " + "{0,-15} {1}" -f "Filter and Sort attributes:", "size, vbr, abr, width, duration, path, parent, ratio, count")
	write-host("  " + "{0,-15} {1}" -f "Output to file in provided:", "-tofile will output the result to a file media_info.txt in the provided folder to be analyzed `n")
}

$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"
$NonExisting_Folder_Error = New-Object System.FormatException "-d provided folder does not exist!"

filter Get-Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

function Get-MeanAttributeResults([string]$folderPath){
	$attrIdx = @{}
	$attrIdx.add("TotalBitrate", 286)
	$attrIdx.add("AudioBitrate", 28)
	$attrIdx.add("FrameWidth", 285)
	$attrIdx.add("Duration", 27)
	
	$shellObject = New-Object -ComObject Shell.Application
	$files = @(Get-ChildItem -literalPath $folderPath -recurse | Get-Where-Extension $t)
	$result = @{}
	foreach($key in $attrIdx.keys){
		$result.add($key, 0)
	}
	
	if($files.count -gt 0){
		foreach($file in $files){
			$directoryObject = $shellObject.NameSpace($file.Directory.FullName)
			$fileObject = $directoryObject.ParseName($file.Name)
			foreach($key in $attrIdx.keys){
				$attrString = $directoryObject.GetDetailsOf($fileObject, $attrIdx[$key])
				
				if($key -eq "Duration"){ 
					$spanRaw = $attrString -Split ":"
					$hours = [int]$spanRaw[0]
					$min = [int]$spanRaw[1]
					$sec = [int]$spanRaw[2]
					$time = [math]::Round(($hours*60 + $min + $sec/60),2)
					$result.$key += $time
				}elseif($attrString -match '\d+') { 
					[int]$attrValue = $matches[0]
					$result.$key += $attrValue
				}else{
					$attrValue = 0
					$result.$key += $attrValue
				}
				
			}
		}
		
		foreach($key in $attrIdx.keys){
			if(-not($key -eq "Duration")){ 
				$result.$key = $result.$key/$files.count
			}
		}
	}
	return $result
}

function Get-MeanAttributeByIndex([string]$folderPath, [int]$attrIndex){
	$shellObject = New-Object -ComObject Shell.Application
	$files = @(Get-ChildItem -literalPath $folderPath -recurse | Get-Where-Extension $t)
	$total = 0
	
	if($files.count -gt 0){
		foreach($file in $files){
			$directoryObject = $shellObject.NameSpace( $file.Directory.FullName )
			$fileObject = $directoryObject.ParseName( $file.Name )
			
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

function Get-MeanAttributeByName([string]$folderPath, [string]$attrName){
	$shellObject = New-Object -ComObject Shell.Application
	$files = @(Get-ChildItem -literalPath $folderPath -recurse | Get-Where-Extension $t)
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

function Get-MeanAttributesValuesData([string]$folderPath, [boolean]$recurse = $r, [System.Data.DataTable]$table){
	$folders = @(Get-ChildItem -literalPath $folderPath | Where-Object{ $_.PSIsContainer })
	$progress = 0
	
	if($folders.count -gt 0){
		foreach($folder in $folders) {
			Write-Progress -Activity "Search folders" -Status "$progress% Complete:" -PercentComplete $progress
			if ($recurse){Get-MeanAttributesValuesData $folder.Fullname $false ([ref]$table)}	
			
			$values = Get-MeanAttributeResults $folder.Fullname
			$row = $table.NewRow()
			$size = Get-ChildItem -literalpath $folder.fullname -r | Measure-Object -property length -sum
			$row["size"] = [math]::Round($size.sum / "1MB")
			$row["vbr"] = $values.TotalBitrate 
			$row["abr"] = $values.AudioBitrate 
			$row["width"] = $values.FrameWidth 
			$row["duration"] = $values.Duration 
			$row["path"] = $folder.Name
			$row["parent"] = (get-item $folder.Fullname ).parent.Name
			$row["count"] = @(Get-ChildItem -literalPath $folder.Fullname -recurse | Get-Where-Extension $t).count
			
			if($row["duration"] -eq 0){
				$row["ratio"] = 0
			}else{
				$row["ratio"] = [math]::Round(($row["size"]/$row["duration"]), 2)
			}
			
			if (RowMatchesFilterCriteria($row)){
				$table.Rows.Add($row)
			}

			$i = $i +1
			$progress = [math]::Round(($i/$folders.count)*100)
		}	
	}else{
		Write-Progress -Activity "Search folder" -Status "$progress% Complete:" -PercentComplete $progress
		$folder = Get-Item -literalPath $folderPath
		$values = Get-MeanAttributeResults $folder.Fullname
		$row = $table.NewRow()
		$size = Get-ChildItem -literalpath $folder.fullname -r | Measure-Object -property length -sum
		$row["size"] = [math]::Round($size.sum / "1MB")
		$row["vbr"] = $values.TotalBitrate 
		$row["abr"] = $values.AudioBitrate 
		$row["width"] = $values.FrameWidth 
		$row["duration"] = $values.Duration 
		$row["path"] = $folder.Name
		$row["parent"] = ""
		$row["count"] = @(Get-ChildItem -literalPath $folder.Fullname -recurse | Get-Where-Extension $t).count
		if($row["duration"] -eq 0){
			$row["ratio"] = 0
		}else{
			$row["ratio"] = [math]::Round(($row["size"]/$row["duration"]), 2)
		}
		
		if (RowMatchesFilterCriteria $row){
			$table.Rows.Add($row)
		}
	}
}

function RowMatchesFilterCriteria([System.Object] $row){
	$output = $true
	$ops = ("<=",">=","<",">","=")
	
	foreach($f in $filter){
		foreach($op in $ops){
			if($f.Contains($op)){
				$splitOp = $op
				break
			}
		}

		$attr = $f.split($splitOp)[0].trim()
		$val = $f.split($splitOp)[1].trim()
		
		if ($row[$attr]){
			$foutput = (IsLogicalConditionMet $row[$attr] $val $splitOp)
			$output = ( $foutput -and $output)
		}
	}
	return $output
}

function IsLogicalConditionMet($val1, $val2, $op){
	if (-not ([string]($val1 -as [int])) -and ([string]($val2 -as [int]))) {
		return $false
	} else {
		$val1 = $val1 -as [double]
		$val2 = $val2 -as [double]
		switch ($op) {
			"<=" { return ($val1 -le $val2) }
			">=" { return ($val1 -ge $val2) }
			"<" { return ($val1 -lt $val2) }
			">" { return ($val1 -gt $val2) }
			"=" { return ($val1 -eq $val2) }
			Default {return ($val1 -eq $val2)}
		}
	}
}
function AdjustConsoleWidthToTableOutput([System.Data.DataTable]$table){
	#getting max row path length
	$rLen = 0
	foreach($r in $table.Rows){
		$currLen = $r["path"].length + $r["parent"].length
		if($rLen -lt $currLen){$rLen = $currLen}
	}
	
	foreach($p in $properties){
		$s = $p['Name']
		$rlength = $rlength + $s.length
	}
	
	$resultWidth = $rlength + $rLen - 13 #13 is the length of path column header
	
	if( $Host -and $Host.UI -and $Host.UI.RawUI ) {
		$rawUI = $Host.UI.RawUI
		$oldSize = $rawUI.BufferSize
		$typeName = $oldSize.GetType().FullName
		$newSize = New-Object $typeName($resultWidth, $oldSize.Height)
		$rawUI.BufferSize = $newSize
	}
}

function Get-MeanAttributeValues([string]$folderPath, [boolean]$recurse, [boolean]$wrapTable){
	$table = New-Object System.Data.DataTable
	$table.Columns.Add("size","string") | Out-Null
	$table.Columns.Add("vbr","int32") | Out-Null
	$table.Columns.Add("abr","int32") | Out-Null
	$table.Columns.Add("width","int32") | Out-Null
	$table.Columns.Add("duration","float") | Out-Null
	$table.Columns.add("path","string") | Out-Null
	$table.Columns.add("parent","string") | Out-Null
	$table.Columns.add("ratio","float") | Out-Null
	$table.Columns.add("count","int32") | Out-Null
	
	Get-MeanAttributesValuesData $folderPath $recurse ([ref]$table)
	
	$properties = @(
		@{Name="MB/Min ";Expression={$_["ratio"]};Alignment='Center'}
		@{Name="Size(MB) ";Expression={$_["size"]};Alignment='Center'}
		@{Name="Length(Min) ";Expression={$_["duration"]};Alignment='Center'}
		@{Name="#Files ";Expression={$_["count"]};Alignment='Center'}
		@{Name="Video BR ";Expression={$_["vbr"]};Alignment='Center'}
		@{Name="Audio BR ";Expression={$_["abr"]};Alignment='Center'}
		@{Name="Frame Width ";Expression={$_["width"]};Alignment='Center'}
		@{Name="Forlder Name ";Expression={$_["path"]};Alignment='Left'}
		@{Name="Forlder Parent ";Expression={$_["parent"]};Alignment='Left'}
	)
	
	try{AdjustConsoleWidthToTableOutput $table}catch{}
	
	$dw = New-Object System.Data.DataView($table)
	if($sort){
		$dw.Sort="$sort ASC"
	}else{
		$dw.Sort="ratio DESC"
	}
	$dw | Format-Table -Property $properties -AutoSize -Wrap:$wrapTable
}

if($help.ispresent){Get-Help-Message;exit}
while(-not $d){$d = $(Read-Host 'Source Folder')}
if(-not $d) {throw $Missing_Folder_Error}
if(-not $t) {throw $Missing_Filetype_Error}
if(-not (Test-Path $d -PathType Container)) {throw $NonExisting_Folder_Error}

if($tofile.ispresent){
	Get-MeanAttributeValues $d $r $wrap | Out-File (Join-Path $d "media_info.txt")
}else{
	Get-MeanAttributeValues $d $r $wrap
}
