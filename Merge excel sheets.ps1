$ExcelObject=New-Object -ComObject excel.application
$ExcelObject.visible=$true
$ExcelFiles=Get-ChildItem -Path C:\Test
$Workbook=$ExcelObject.Workbooks.add()
$Worksheet=$Workbook.Sheets.Item("Sheet1")
$outFile = "C:\Test\data.txt"
foreach($ExcelFile in $ExcelFiles){
 
$Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)
foreach($worksheet in $Everyexcel.sheets){
 $endRow = $worksheet.UsedRange.SpecialCells(11).Row

        $rangeAddress = $worksheet.Cells.Item(2, 1).Address() + ":" + $worksheet.Cells.Item($endRow, 1).Address()
        Write-Host "Using range $($rangeAddress)"
        $worksheet.Range($rangeAddress).Value2 | Out-File -FilePath $outFile -Append
    #    $workbook.Close($false) 
 }
#$Everysheet=$Everyexcel.sheets.item(1)
#$Everysheet.Copy($Worksheet)
#$Everyexcel.Close()
 
}
#$Workbook.SaveAs("C:\Test\merge.xlsx")
$ExcelObject.Quit()