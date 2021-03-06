 param (
	[string]$f
 )

enum AttributeIndex 
{
    TotalBitrate = 286
    AudioBitrate = 28
    FrameWidth = 285
    Duration = 27
}

function Get-FileAttributeValueById($file, [int]$attrIndex){
    if($file.gettype().name -eq "String"){
        $file = (Get-Item $file)
    }
	$shellObject = New-Object -ComObject Shell.Application
	$directoryObject = $shellObject.NameSpace($file.Directory.FullName)
	$fileObject = $directoryObject.ParseName($file.Name)
	$attrString = $directoryObject.GetDetailsOf( $fileObject, $attrIndex )
    if( $attrString -match '\d+' ) { 
        [int]$attrValue = $matches[0] 
    }else { 
        $attrValue = -1 
    }
	return $attrValue
}

function Get-AudioQualityToEncode($file){
	#set audio bitrate to a min of the original or max of provided
	return (Get-FileAttributeValueById $file ([System.Int32][AttributeIndex]::AudioBitrate))
}

#Check if source dir was provided -> ask for source
while(-not $f){$f = $(Read-Host 'File Path')}
if(!(Test-Path -literalPath $f)) {throw $NonExisting_Folder_Error}
Get-AudioQualityToEncode $f
