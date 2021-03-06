 param (
	[string]$src,
	[string]$dst,
	[String[]]$type = ('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv', '.mpg', '.ts'),
	[String]$action = 'move',
	[switch]$rename = $false,
	[switch]$r = $true,
	[switch]$help
 )
 
$Missing_Filetype_Error = New-Object System.FormatException "-type (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-src (source folder) is missing!"

function Get-Help-Info{
    write-host("`nSYNTAX")
	write-host("  " + "Copy-Files.ps1 [-src]<string> [[-dst]<string>] [[-type]<string[]=('.avi', '.mp4', '.wmv', '.flv', '.mov', '.m4v', '.f4v', '.mkv', '.mpg', '.ts')>] [[-action]<string>='move'|'copy'] [[-rename]<switch>] [[-r]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-src", "Provides the dir path to be searched for files to copy.")
	write-host("  " + "{0,-15} {1}" -f "-dst", "Provides the dir path to be searched for media files to convert.")
	write-host("  " + "{0,-15} {1}" -f "-type", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-action", "move or copy")
	write-host("  " + "{0,-15} {1}" -f "-rename", "Destination files will inherit the name of their parent folders")
	write-host("  " + "{0,-15} {1}" -f "-r", "Search directory recursively for files.")
	exit
}

filter Get-Files-Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

function Move-Files(){
	if (-not $dst){$dst = $src}
	
	#get file recursively if -Recurse was provided
	
	if ($r.IsPresent){
		$files = @(Get-ChildItem -literalPath $src -recurse | Get-Files-Where-Extension $type)
	}else{
		$files = @(Get-ChildItem -literalPath $src | Get-Files-Where-Extension $type)
	}
	
	if ($files.count -eq 0){
		write-host "No files with filter: " $type " were found in " $src
		exit
	}
	
	write-host "============================================================" 
	write-host "Found " $files.count " Files"
	write-host "============================================================"
	
	$dupCounter = 1
	foreach ($file in $files){
		if ($rename){
			$srcFileDir = (get-item -path $src).FullName
			$currFileDir = (get-item -path $file.fullname).directory.FullName
			$prevFileDir = $currFileDir
			
			while(($srcFileDir -ne $currFileDir) -and ($currFileDir.length -ne 0)){
				$prevFileDir = $currFileDir
				$currFileDir = (get-item -path $currFileDir).parent.FullName
			}
			
			if($prevFileDir -eq $srcFileDir){
				$fileName = $file.Name
			}else{
				$fileName = (get-item -path $prevFileDir).Name + $file.extension
			}
		}else{
			$fileName = $file.Name
		}
		
		$dstFilePath = join-path -path $dst $fileName
		
		if (Test-Path $dstFilePath -PathType Leaf){ 
			$newFileName = (get-item -path $dstFilePath).basename + "_" + $dupCounter.tostring() + $file.extension
			$dstFilePath = join-path -path $dst $newFileName
			$dupCounter++
		}
		
		if ($action -eq "move"){
			$actionPrint = "Moved"
			Move-Item -literalPath $file.FullName -Destination $dstFilePath -force
		}elseif ($action -eq "copy"){
			$actionPrint = "Copies"
			Copy-Item -literalPath $file.FullName -Destination $dstFilePath -force
		}
		
		write-host $actionPrint `t $file.FullName " ---> " 
		write-host `t $dstFilePath
		write-host "-------------------------------------------------------------"
	}
}

################## BEGIN EXECUTION ##################

#Check if help was invoked
if($help.ispresent){Get-Help-Info}

#Check if source dir was provided -> ask for source
while(-not $src){$src = $(Read-Host 'Source Folder')}
if(!(Test-Path $src -PathType Container)) {throw $NonExisting_Folder_Error}

#Move files matching the types
Move-Files
