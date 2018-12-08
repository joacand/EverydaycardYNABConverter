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
$dt.Columns.Add("Inflow")
$et = $dt | select -Skip 5

$result = @()

foreach ($i in $et) {
    If ([int]$i[4] -lt 0) { continue }
    $i[0] = ([datetime]($i[0])).ToString('yyyy/MM/dd')
    $i[1] = $i[3]
    $i[2] = ''
    $i[3] = ''
    $i[4] = $i[4] -replace '[,]', '.'
    $i[5] = ''
    $result += $i
}

$result | ConvertTo-Csv -NoTypeInformation -Delimiter "," | % {$_ -replace '"', ''} | Out-File $outputFile -encoding utf8
