# Parameters
Param(
    [Parameter(Mandatory=$true)][String]$Source,
    [String]$TableName
)

If(!$TableName){
    $TableName = "Table"
}

$source = (Get-Item $source).FullName
$destination = $source.Replace(".csv",".xlsx")

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$workbook = $Excel.Workbooks.Open($source)
$worksheet = $workbook.ActiveSheet

$ListObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.Name = $TableName

$workbook.SaveAs($destination,51)
$workbook.Close()
$Excel.quit()