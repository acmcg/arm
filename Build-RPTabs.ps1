$excelIDs = Get-Process Excel -ea SilentlyContinue #work around until I can figure out how to clean up Objects
$TemplateDocsColumn = 9
$ProviderNamespaceColumn = 3
$ResourceColumn = 7
$location = Get-Location 
$dateRegEx = "([2]\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01]))"
$excelObject=new-object -com excel.application
$workBook=$excelObject.workbooks.open("$location\ISM Protected Controls.xlsx")
$excelObject.visible = $True
$summarySheet = $Workbook.Sheets.Item("IRAP Service Summary")
$apiDocVersionArray = @()
$rowMax = ($summarySheet.UsedRange.Rows).count
for ($row=2; $row -le $rowMax; $row++){ # start at 2 to ignore the column header
    if($summarySheet.Cells.Item($row, $TemplateDocsColumn).text -match $dateRegEx){
        $apiDocVersionObject = @{Namespace = $($summarySheet.Cells.Item($row, $ProviderNamespaceColumn).text);`
                                    Resource = $($summarySheet.Cells.Item($row, $ResourceColumn).text);`
                                    ApiVersion = $($summarySheet.Cells.Item($row, $TemplateDocsColumn).text)}
        $apiDocVersionArray += $apiDocVersionObject
        }
}
$excelObject.workbooks.Close() #$false - doesn't save changes
$excelObject.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject)
#Create Excel Workbook
Remove-Variable excelObject
$excelObject = New-Object -ComObject excel.application 
$outputpath = Get-Location
$workBook=$excelObject.workbooks.open("$location\ISM Protected Controls.xlsx")
$excelObject.visible = $True

foreach ($element in $apiDocVersionArray){
    $apiVersion = $element.ApiVersion
    $resource = $element.Resource
    $resourceProvider = $element.Namespace
    $url = "https://docs.microsoft.com/en-au/azure/templates/$resourceProvider/$apiVersion/$resource"
    $html = Invoke-WebRequest $url -usebasicparsing
    $htmlArray = $html.Content -split"<pre><code class=`"lang-json`">|</code></pre>"
    $json = $htmlArray[1]
    $json = $json -replace "&quot;","`""
    $json = $json -replace "boolean","`"boolean`""
    $sheetExists = $null  # Even though $workbook.worksheets.item($sheetName) fails it doesn't set the variable to NULL
    if(!(test-path -path ".\$resourceProvider")){
        New-Item -ItemType Directory -Force -Path ".\$resourceProvider"
    }
    #remove path like structure from Resource
    $resource= $resource -replace "/","-"
    $json | Out-File -FilePath ".\$resourceProvider\$resource-$apiVersion.json" -force
    # # Create workbooks
    $sheetName = $resourceProvider.Split(".")[1]
    $sheetExists = $workbook.worksheets.item($sheetName)
    if($sheetExists){ # If sheet already exists
        $currentSheet = $workbook.worksheets.item($sheetName)
        $rowMax = ($currentSheet.UsedRange.Rows).count
        $startingExcelRow = $rowMax +5 # write the JSON data after the last row
    }
    else{
        $currentSheet = $workbook.Worksheets.add()
        $startingExcelRow = 2 # write the JSON two rows from the top of the sheet
        $currentSheet.Name = $sheetName
        # Set header info
        $currentSheet.Cells.item(1,2) = "Australian Central Protected"
        $currentSheet.Cells.item(2,2) = "Must"
        $currentSheet.Cells.item(2,3) = "Should"
        $currentSheet.Cells.item(2,4) = "Control"
        $currentSheet.Cells.item(2,5) = "Description"
        $currentSheet.Cells.item(1,7) = "Australian Protected"
        $currentSheet.Cells.item(2,7) = "Must"
        $currentSheet.Cells.item(2,8) = "Should"
        $currentSheet.Cells.item(2,9) = "Control"
        $currentSheet.Cells.item(2,10) = "Description"
    }
    $content = get-content -raw ".\$resourceProvider\$resource-$apiVersion.json"
    $contentArray = $content.Split([Environment]::NewLine)
    for($arrayIndex = 0 ; $arrayIndex -lt $contentArray.Length ; $arrayIndex++){
        $currentSheet.Cells.item($startingExcelRow+$arrayIndex+1,1) = $contentArray[$arrayIndex] # add 1 to arrayindex because excel does have a row 0
    }
    $output = $currentSheet.Columns.AutoFit()  


}
#$currentSheet = $workbook.Worksheets.add()
$fileNameSuffix = get-date -Format yyMMddHHmm
#$excelObject.ActiveWorkbook.SaveAs("$outputpath\ISM-$fileNameSuffix.xlsx") 
$excelObject.ActiveWorkbook.SaveAs("$location\ISM Protected Controls.xlsx")
$excelObject.Workbooks.Close()
$excelObject.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject)

$leftoverExelObj = Get-Process Excel -ea SilentlyContinue | ? {$_.Id -notin $excelIDs.Id}
if ($leftoverExelObj)
{
    $leftoverExcelObj.Kill() | Out-Null
} 


