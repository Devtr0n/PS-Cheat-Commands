<# 
	Author: Richard Hollon - www.richardhollon.com/Contact.aspx
	This script downloads an MS Word document and replaces targeted text values with user input data.
#>
#======================================================================================================================================

#1.) Get User Input
$targetLabel		= Read-Host 'Enter the target release label: '
$previousLabel		= Read-Host 'Enter the previous release label: '
$productionLabel	= Read-Host 'Enter the production release label: '
$releaseDate 		= Read-Host 'Enter the target release date: '
trap {'You did not enter a valid target date!'; continue}
. {
  $releaseDate = [DateTime]::Parse($releaseDate)
}

#2.) Initialize all variables
$findText 			= @("{targetLabel}","{previousLabel}","{productionLabel}","{releaseDate}")
$userInput			= @($targetLabel,$previousLabel,$productionLabel,$releaseDate)
$templateFileURL	= 'http://myserver/templates/Template.MyProject.PromotionRequest.Test.docx'
$outFilePath		= [Environment]::GetFolderPath("Desktop")+'\MyProject.PromotionRequest.Test.docx'

#3.) Open CEMS Sharepoint template document via Word Interop
$word 				= New-Object -com word.application
$document 			= $word.Documents.Open($templateFileURL)
$word.Visible		= $False #make this true to watch Word process changes in real-time! 
#Write-Host $document.Range().text #this is all the text in the word doc!

#4.) Replace target text in Word document (make the desired changes).
foreach ($item in $findText) {
	Replace-Word-Text -FindText $item -ReplaceText $userInput[$findText.IndexOf($item)]
}

#5.) Close MS-Word and Document
$document.Close()
$word.Quit()

#======================================================================================================================================
function Replace-Word-Text(
[string]$FindText,
[string]$ReplaceText
){
    $ReplaceAll 		= 2
    $FindContinue 		= 1
    $MatchCase 			= $False
    $MatchWholeWord 	= $True
    $MatchWildcards 	= $False
    $MatchSoundsLike 	= $False
    $MatchAllWordForms 	= $False
    $Forward 			= $True
    $Wrap 				= 1
    $Format 			= $False

    $word.Selection.Find.Execute(
        $FindText,
        $MatchCase,
        $MatchWholeWord,
        $MatchWildcards,
        $MatchSoundsLike,
        $MatchAllWordForms,
        $Forward,$Wrap,$Format,
        $ReplaceText,
        $ReplaceAll) 

    $document.SaveAs([ref]$outFilePath) #Save to user's Desktop
}
#======================================================================================================================================