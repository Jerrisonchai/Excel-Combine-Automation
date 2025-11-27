# 📂 Excel Combine Automation
Automate merging multiple Excel `.xlsx` files into a single combined file using a PowerShell script, triggered by a `.bat` launcher. Useful for consolidating reporting files, daily exports, or multi-department contributions without manual copy-paste.

---

## 🚀 Features
- Combine all Excel files inside a selected folder
- Automatically copies headers only once
- Appends data rows from each file sequentially
- Uses COM automation (no external libraries required)
- Optional `.bat` file allows 1-click execution
- No need to open Microsoft Excel manually

---

## 📁 Folder Structure
|-- combine_excel.ps1 # PowerShell script that merges files
|-- run.bat # Batch file to execute script with bypass policy


---

## 🛠 Requirements
| Tool | Minimum Requirement |
|-------|---------------------|
| Windows | Windows 10 or above |
| Excel | Microsoft Excel installed |
| Execution Permission | Script execution allowed (`Bypass` via BAT included) |

---

## 🧠 How It Works
1. PowerShell opens Excel through COM automation
2. Creates a new workbook called `Combined.xlsx`
3. Copies:
   - Header row (A1) only once
   - Remaining rows (A2 to last row) from each file sequentially
4. Saves the result to the defined output path

---

## 📌 PowerShell Script (`combine_excel.ps1`)
```powershell
$sourceFolder = "C:\Users\Jerrison Chai\Documents\02 DEMO\TempRawData"
$outputFile   = "C:\Users\Jerrison Chai\Documents\02 DEMO\TempRawData\Combined.xlsx"

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
```

## Run Using BAT Launcher (run.bat)
```
@echo off
PowerShell -ExecutionPolicy Bypass -File "C:\Users\Jerrison Chai\Documents\02 DEMO\combine_excel.ps1"
pause
```
## 📥 Usage Instructions
- Put all .xlsx files into the folder specified in $sourceFolder
- Update file paths as needed inside both ps1 and bat files
- Double-click run.bat
- Wait for completion message
- Open Combined.xlsx to review result

## ⚠ Notes
- All Excel files must have the same column structure
- Script reads Sheet 1 only by default (similar sheet across files)
- Does not support password-protected files (extended version available on request)

## 💡 Future Enhancements
- Add sheet name prompt to merge specific sheets
- Add support for password-protected Excel files
- Auto-generate timestamped filenames
