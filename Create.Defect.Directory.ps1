<# 
	Author: Richard Hollon
	This script creates and initializes a new directory for a developer's defect notes.
#>
#=================================================================================================================================
#1.) Get User Input
$defectNumber				= Read-Host 'Enter the defect number: (ex. defect1308) ' 
if ($defectNumber -notcontains '*defect*') { 
	$defectNumber 			= 'defect' + $defectNumber
}
$outPath               		= [Environment]::GetFolderPath("MyDocuments")+'\defects\'+$defectNumber+'\'

#2.) Create a directory for defect's notes
if((Test-Path $outPath) -eq 0)
{
	md $outPath #make directory
} else {
	Write-Host 'Directory already exists.'
}

#3.) Create empty text file for defect's notes
New-Item 	($outPath + $defectNumber + '-notepad.txt') -type file
Write-Host 	($outPath + $defectNumber + '-notepad.txt') ' created!!'
#=================================================================================================================================