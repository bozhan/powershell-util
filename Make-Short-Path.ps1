#Shortpath
$a = New-Object -ComObject Scripting.FileSystemObject
$f = $a.GetFile("C:\Scripttest\LongFileNameTest.txt")
$f.ShortPath