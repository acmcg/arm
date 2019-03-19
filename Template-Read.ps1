
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
    
    if(!(test-path -path ".\$resourceProvider")){
        New-Item -ItemType Directory -Force -Path ".\$resourceProvider"
    }
    #remove path like structure from Resource
    $resource= $resource -replace "/","-"
    $json | Out-File -FilePath ".\$resourceProvider\$resource-$apiVersion.json" -force
}

$excelObject.Documents.Close()

