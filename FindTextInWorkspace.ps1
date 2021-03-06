<# 
Author: Richard Hollon - www.richardhollon.com/Contact.aspx
This script searches a directory for a specific string pattern. Limitation is that it does not search folder names
#>
$pattern='spMyStoredProcedureName'
$directory='C:\Workspace\somefolder\*'
$list=@(dir $directory -recurse | Get-ChildItem | Select-String -pattern $pattern | Select-Object Path -Unique)
$list+=@(dir $directory -recurse | Get-ChildItem | Where-Object { $_.Name -like '*' + $pattern + '*' } | Select-Object Fullname) 
$list.Count
$list | % { $_.Path;$_.Fullname }
