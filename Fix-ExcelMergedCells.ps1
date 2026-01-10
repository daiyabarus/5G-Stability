# ============================================================================
# Fix-ExcelMergedCells.ps1
# Post-processing script to fix merged cells in Excel report
# ============================================================================

<#
.SYNOPSIS
    Fixes merged cells in Performance sheets after Python processing.

.DESCRIPTION
    This script opens the generated Excel report and ensures all tower blocks
    have properly merged cells matching the template format.

.PARAMETER ExcelPath
    Path to the Excel file to fix

.EXAMPLE
    .\Fix-ExcelMergedCells.ps1 -ExcelPath "C:\Reports\5G_Detail_Report_20260110.xlsx"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath
)

# Validate file exists
if (-not (Test-Path $ExcelPath)) {
    Write-Error "File not found: $ExcelPath"
    exit 1
}

Write-Host "Starting Excel merged cells fix..." -ForegroundColor Cyan
Write-Host "Target file: $ExcelPath" -ForegroundColor Gray

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open workbook
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    Write-Host "`nProcessing sheets..." -ForegroundColor Yellow
    
    # Process each worksheet
    foreach ($sheet in $workbook.Worksheets) {
        # Only process "Performance" sheets
        if ($sheet.Name -like "Performance*") {
            Write-Host "  Processing: $($sheet.Name)" -ForegroundColor Green
            
            $currentRow = 1
            $blockCount = 0
            
            # Process blocks (each block is 13 rows)
            while ($true) {
                # Check if Tower ID cell (B1 + offset) has content
                $towerIdCell = $sheet.Cells.Item($currentRow, 2)
                
                if ([string]::IsNullOrWhiteSpace($towerIdCell.Text)) {
                    break  # No more tower blocks
                }
                
                $blockCount++
                Write-Host "    Block $blockCount at row $currentRow" -ForegroundColor DarkGray
                
                # Define merged cell ranges for this block
                $merges = @(
                    # B1:C1 - New Tower ID label
                    @{Start=@($currentRow, 2); End=@($currentRow, 3)},
                    
                    # B2:C2 - Batch label
                    @{Start=@($currentRow+1, 2); End=@($currentRow+1, 3)},
                    
                    # B3:C3 - START-TIME label
                    @{Start=@($currentRow+2, 2); End=@($currentRow+2, 3)},
                    
                    # B4:C4 - Date label
                    @{Start=@($currentRow+3, 2); End=@($currentRow+3, 3)},
                    
                    # B7:E7 - Header row (if exists)
                    @{Start=@($currentRow+6, 2); End=@($currentRow+6, 5)}
                )
                
                # Apply merges
                foreach ($merge in $merges) {
                    $startRow = $merge.Start[0]
                    $startCol = $merge.Start[1]
                    $endRow = $merge.End[0]
                    $endCol = $merge.End[1]
                    
                    try {
                        $range = $sheet.Range(
                            $sheet.Cells.Item($startRow, $startCol),
                            $sheet.Cells.Item($endRow, $endCol)
                        )
                        
                        # Unmerge first (in case partially merged)
                        try { $range.UnMerge() } catch {}
                        
                        # Merge cells
                        $range.Merge()
                        
                    } catch {
                        Write-Warning "      Failed to merge range at row $startRow"
                    }
                }
                
                # Move to next block (13 rows + 2 spacing)
                $currentRow += 15
            }
            
            Write-Host "    Total blocks processed: $blockCount" -ForegroundColor Cyan
        }
    }
    
    # Save and close
    Write-Host "`nSaving changes..." -ForegroundColor Yellow
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "`n✓ Excel file fixed successfully!" -ForegroundColor Green
    Write-Host "File: $ExcelPath" -ForegroundColor Gray
    
} catch {
    Write-Error "Error processing Excel file: $_"
    
    # Cleanup on error
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    
    exit 1
}
