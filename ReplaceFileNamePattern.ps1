dir 'C:\Workspace\Project\' | gci | rename-item -newname { $_.name -replace 'Bert','Ernie' }
