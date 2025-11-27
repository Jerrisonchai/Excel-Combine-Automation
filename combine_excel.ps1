$sourceFolder = "C:\...\TempRawData"
$outputFile   = "C:\...\TempRawData\Combined.xlsx"

Write-Host "Starting Excel file combination..." -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Create output workbook
$combinedWb = $excel.Workbooks.Add()
$combinedSheet = $combinedWb.Sheets.Item(1)
$rowOffset = 1

Get-ChildItem "$sourceFolder\*.xlsx" | ForEach-Object {
    Write-Host "Processing file: $($_.Name)" -ForegroundColor Yellow  # Show progress

    $wb = $excel.Workbooks.Open($_.FullName)
    $ws = $wb.Sheets.Item(1)

    $lastRow = $ws.UsedRange.Rows.Count

    if ($rowOffset -eq 1) {
        $ws.Range("A1").EntireRow.Copy($combinedSheet.Range("A$rowOffset"))
        $rowOffset++
    }

    $ws.Range("A2:A$lastRow").EntireRow.Copy($combinedSheet.Range("A$rowOffset"))
    $rowOffset += ($lastRow - 1)

    $wb.Close($false) | Out-Null   # Prevents "True" output
}

# Save
$combinedWb.SaveAs($outputFile)
$combinedWb.Close()
$excel.Quit()

Write-Host "All files combined successfully!" -ForegroundColor Green
Write-Host "Output saved to: $outputFile" -ForegroundColor White
