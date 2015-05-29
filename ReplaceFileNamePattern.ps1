<# 
Author: Richard Hollon - www.richardhollon.com/Contact.aspx

This script renames items in a directory by replacing a string pattern with a new string value.

#>
dir 'C:\Workspace\Project\' | gci | rename-item -newname { $_.name -replace 'Bert','Ernie' }
