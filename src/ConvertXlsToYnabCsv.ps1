param (
    [Parameter(Mandatory = $true)][string]$inputFile,
    [Parameter(Mandatory = $true)][string]$outputFile
)

$strFileName = $inputFile
$strSheetName = 'Sheet0$'
$strProvider = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSource = "Data Source = $strFileName"
$strExtend = "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';"
$strQuery = "Select * from [$strSheetName]"

$objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
$sqlCommand.Connection = $objConn
$objConn.open()

$da = New-Object system.Data.OleDb.OleDbDataAdapter($sqlCommand)
$dt = New-Object system.Data.datatable
[void]$da.fill($dt)

$objConn.close()

$dt.Columns[0].ColumnName = "Date"
$dt.Columns[1].ColumnName = "Payee"
$dt.Columns[2].ColumnName = "Category"
$dt.Columns[3].ColumnName = "Memo"
$dt.Columns[4].ColumnName = "Outflow"
$dt.Columns[5].ColumnName = "Inflow"
$dt.Columns.RemoveAt(6) # Moms - not used
$et = $dt | Select-Object -Skip 5 # First five lines are title, description, etc.

$result = @()

foreach ($i in $et) {
    $i[0] = ([datetime]($i[0])).ToString('yyyy/MM/dd')
    $i[1] = $i[4]
    $i[2] = ''
    $i[3] = ''
    $i[4] = $i[5] -replace '[,]', '.'
    $i[5] = ''
    $result += $i
}

$result | ConvertTo-Csv -NoTypeInformation -Delimiter "," | ForEach-Object { $_ -replace '"', '' } | Out-File $outputFile -encoding utf8
