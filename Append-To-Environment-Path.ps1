 param (
    [string]$a,
		[switch]$o,
		[switch]$u,
		[switch]$help
 )
$Missing_Parameter_Error = New-Object System.FormatException "-a parameter is missing!"

if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Append-To-Environment-Path [[-a]<string>] [[-u]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-a", "Specify the string to be appended to the path environment variable")
	write-host("  " + "{0,-15} {1}" -f "-o", "Overwrites the PATH variable with supplied string in -a param")
	write-host("  " + "{0,-15} {1}" -f "-u", "If present sets the USER leve path variable, else MACHINE level")
	exit
}

if (-not $a){$a = $(Read-Host 'Source Folder')}
if (-not $a){throw $Missing_Parameter_Error}

if ($u){
	$param = "User"
}else{
	$param = "Machine"
}

$PATH = [Environment]::GetEnvironmentVariable("PATH", $param)
write-host("`nBefore edit:")
write-host([Environment]::GetEnvironmentVariable("PATH", $param))

if ($o){
	[Environment]::SetEnvironmentVariable("PATH", "$a;", $param)
	write-host("`nAfter edit:")
	write-host([Environment]::GetEnvironmentVariable("PATH", $param))
	exit
}

if( $PATH -notlike "*"+$a+"*" ){
		if ($PATH.substring($PATH.length - 1, 1) -notlike ";"){
				[Environment]::SetEnvironmentVariable("PATH", "$PATH;$a;", $param)
		}else{
			[Environment]::SetEnvironmentVariable("PATH", "$PATH$a;", $param)
		}
		
	write-host("`nAfter edit:")
	write-host([Environment]::GetEnvironmentVariable("PATH", $param))
}