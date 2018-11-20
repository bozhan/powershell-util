#analyzes the size of each subfolder in provided folder and returns the ones with grater size than the provided size limit
param (
	[string]$d,
	[string]$limit = 0,
	[string]$scale = "GB",
	[switch]$help
)

$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"

if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Get-Subfolder-Size [[-d]<string>] [[-limit]<integer=0>], [[-sizeScale<string>=GB]]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Specify directrory analyze.")
	write-host("  " + "{0,-15} {1}" -f "-limit", "Set maximum folder size in MB.")
	exit
}

$d = $(Read-Host 'Source Folder')
if (-not $d){throw $Missing_Folder_Error}
$ds = @(Get-ChildItem -literalpath $d | Where-Object{($_.PSIsContainer)})
$dev = "1" + $scale
write-host $ds.count " sub-dirs found!"
foreach ($dir in $ds){
	$size = Get-ChildItem -literalpath $dir.fullname -r | Measure-Object -property length -sum
	
	if (($size.sum / $dev) -gt $limit){
		write-host ("{0:N2}" -f ($size.sum / $dev) + " $scale"  + " : " + $dir.name)
	}
}

