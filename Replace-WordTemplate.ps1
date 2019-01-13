<#

.SYNOPSIS
This is a simple Powershell script to replace UniqueID's in a Word Template with Data

.DESCRIPTION
The script will use 

.EXAMPLE


.NOTES
  Author:   Leigh Butterworth
  Version:  1.0

.LINK
https://github.com/L37hal/

#>

Param(
    [parameter(Mandatory=$false)][string]$Template,
    [parameter(Mandatory=$false)][string]$OutPath,
    [parameter(Mandatory=$false)][array]$Mappings,
    [parameter(Mandatory=$false)][array]$Dataset
) # End Param()

# *** Entry Point to Functions ***

# *** Entry Point to Script ***

$MatchCase = $False 
$MatchWholeWord = $true
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$wdFindContinue = 1 
$Wrap = $wdFindContinue 
$Format = $False 
$wdReplaceNone = 0 
$wdReplaceAll = 2

$Headers = $Mappings | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'

$Word = New-Object -comobject Word.Application  
$Word.Visible = $False

$Document = $Word.Documents.Open($Template) 
$DocumentText = $Word.Selection 

ForEach ($Header in $Headers)
{
    $UniqueID = $Mappings.$Header | Out-String
    $Value = $Dataset.$Header | Out-String

    [string]$FindText = $UniqueID.Trim()
    [string]$ReplaceWith = $Value.Trim()

    $replace = $DocumentText.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$ReplaceWith,$wdReplaceAll)
}


$Document.SaveAs($OutPath)
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
Remove-Variable Word