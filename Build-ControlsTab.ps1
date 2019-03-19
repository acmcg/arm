# Usage ./ISM-Read.ps1 | convertto-csv -notypeinformation | out-file controls.csv

$Word = New-Object -ComObject Word.Application
$path = get-location
$Document = $Word.Documents.Open("$path\Australian_Government_Information_Security_Manual.docx")
$Word.Visible = $False
$controlObjects = @()
foreach ($paragraphs in $Document.Paragraphs) 
{ 
    if($paragraphs.Style.NameLocal -eq "Heading 1"){
        $ismSection = $paragraphs.range.Text
    }
    # Select only the text 
    If ($paragraphs.Range.Font.ColorIndex -eq 15) 
    { 
       if ($paragraphs.Range.Text -match "Security Control:"){
            $paragraphText = $paragraphs.Range.Text -replace "\v",";" #Remove vertical tab 
            $paragraphArray = $paragraphText -Split "Security Control: |; Revision: |; Updated: |; Applicability: |; Priority: |;" # Split string and remove headings
            $paragraphArray = $paragraphArray.ForEach("Trim") # remove leading and trailng whitespaces
            # start array from position 1 as position 0 is empty due to the way split works
            $controlObjects += [pscustomobject]@{ISMSection = $ismSection;Control = $paragraphArray[1];Revision = $paragraphArray[2];Updated = $paragraphArray[3];Classification = $paragraphArray[4];Priority = $paragraphArray[5];Description = $paragraphArray[6] }
        }
        else{
            #loop on paragraphs as word list items are individual paragraphs
            $controlObjects[-1].Description = "$($controlObjects[-1].Description) `n -$($paragraphs.Range.Text)"
        }
    } 

} 
$controlObjects | convertto-csv -notypeinformation | out-file controls.csv

$Word.Documents.Close()


