
$TemplateDocsColumn = 9
$ProviderNamespaceColumn = 3
$ResourceColumn = 7
$dateRegEx = "([2]\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01]))"
$excelObject=new-object -com excel.application
$workBook=$excelObject.workbooks.open("C:\Users\Andrew.McGregor\arm\ISM Protected Controls.xlsx")
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

#Create Excel Workbook
$excel = New-Object -ComObject excel.application 
$excel.visible = $True
$outputpath = Get-Location

$workbook = $excel.Workbooks.Add()

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
        $startingExcelRow = 0 # write the JSON data at the top of the sheet
        $currentSheet.Name = $sheetName
    }
    $content = get-content -raw ".\$resourceProvider\$resource-$apiVersion.json"
    $contentArray = $content.Split([Environment]::NewLine)
    for($arrayIndex = 0 ; $arrayIndex -lt $contentArray.Length ; $arrayIndex++){
        $currentSheet.Cells.item($startingExcelRow+$arrayIndex+1,1) = $contentArray[$arrayIndex] # add 1 to arrayindex because excel does have a row 0
    }


}
$currentSheet = $workbook.Worksheets.add()
$currentSheet.Name = "ISM Controls"
$workbook.SaveAs("$outputpath\rp.xlsx") 

$excelObject.Documents.Close()

