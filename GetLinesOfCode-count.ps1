<# 
Author: Richard Hollon - www.richardhollon.com/Contact.aspx

#this script measures lines of code, excluding empty lines and comments.
#last write time is greater than 02/01/2015
#>

$dir = 'C:\Workspace\mySolution\myProject\myFolder'
(dir $dir -include *.* -recurse | get-childitem –recurse | where-object {$_.lastwritetime -gt “2/01/2015”} | select-string "^(\s*)//" -notMatch | select-string "^(\s*)$" -notMatch).Count
