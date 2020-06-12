####################################
#         Author:c0rekt            #
#          version 1.1             #
#    Excel to csv converter        #
####################################

function ExcelToCsv {
    [cmdletbinding()]
    Param (
        [string]$ExcelFileName,
        [string]$ExcelFilePath,
        [string]$SheetName,
        [string]$CsvLocation,
        [string]$Delimiter,
        [int]$RowStart
    )
    if(($ExcelFileName) -and ($ExcelFilePath) -and ($SheetName) -and ($CsvLocation) -and ($Delimiter) -and ($RowStart)){
    $excelfile = "$($ExcelFilePath)$($ExcelFileName)"
    $excel = New-Object -ComObject Excel.Application
    $excelWorkbook = $excel.Workbooks.Open("$excelfile")
    $excelWorksheet = $excelWorkbook.Sheets.Item("$SheetName")
    $WorksheetRange = $excelWorksheet.UsedRange
    $RowCount = $WorksheetRange.Rows.Count
    $ColumnCount = $WorksheetRange.Columns.Count

    for($i = $RowStart; $i -le $RowCount; $i++){
        for($j = 1; $j -le $ColumnCount; $j++){
            $items =  $excelWorksheet.Columns.Item($j).Rows.Item($i).Text
                if(!$items){
                    $items = "NULL"
                }
                $output += "$($items)$($Delimiter)"
            }
            $StringLength = $output.Length
            $output = $output.Substring(0,$StringLength-1) + "`n"
    }
    $OutputFileName = [io.path]::GetFileNameWithoutExtension($ExcelFileName)
    Write-output "$($output)" >> "$($CsvLocation)$($OutputFileName).csv"
    }
    else {
        Write-output "Missing parameters, please check again"
    }
}
$excel.Workbooks.Close()
Export-ModuleMember -Function * -Alias *
