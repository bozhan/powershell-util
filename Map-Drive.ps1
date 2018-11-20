New-PSDrive -name "tmp" -psprovider FileSystem -root $dir
#$dirs = @(Get-ChildItem -literalpath "tmp:\" | Where-Object{($_.PSIsContainer)})
#o = join-path -path "tmp:\" $d.name
#$size = Get-ChildItem -literalpath $o -r | Measure-Object -property length -sum
Remove-PSDrive "tmp"