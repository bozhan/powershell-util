param (
    $File,
    [string]$AttributeName
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

Get-FileAttributeValueById $file ([System.Int32][AttributeIndex]::$AttributeName)