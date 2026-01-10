# ============================================================================
# utils/excel_post_processor.py
# ============================================================================
"""Post-processing utilities for Excel files using PowerShell."""

import subprocess
import sys
from pathlib import Path
from typing import Optional


class ExcelPostProcessor:
    """Handle Excel post-processing using PowerShell for merged cells."""
    
    POWERSHELL_SCRIPT = """
# Fix merged cells in Excel Performance sheets
param([string]$ExcelPath)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Name -like "Performance*") {
            $currentRow = 1
            
            while ($true) {
                $towerIdCell = $sheet.Cells.Item($currentRow, 2)
                if ([string]::IsNullOrWhiteSpace($towerIdCell.Text)) { break }
                
                # Merge B1:C1, B2:C2, B3:C3, B4:C4, B7:E7
                $merges = @(
                    @($currentRow, 2, $currentRow, 3),
                    @($currentRow+1, 2, $currentRow+1, 3),
                    @($currentRow+2, 2, $currentRow+2, 3),
                    @($currentRow+3, 2, $currentRow+3, 3),
                    @($currentRow+6, 2, $currentRow+6, 5)
                )
                
                foreach ($m in $merges) {
                    try {
                        $range = $sheet.Range($sheet.Cells.Item($m[0],$m[1]), $sheet.Cells.Item($m[2],$m[3]))
                        try { $range.UnMerge() } catch {}
                        $range.Merge()
                    } catch {}
                }
                
                $currentRow += 15
            }
        }
    }
    
    $workbook.Save()
    $workbook.Close()
    Write-Host "SUCCESS"
} catch {
    Write-Host "ERROR: $_"
    exit 1
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
"""
    
    @staticmethod
    def is_windows() -> bool:
        """Check if running on Windows."""
        return sys.platform == "win32"
    
    @staticmethod
    def fix_merged_cells(excel_path: str) -> tuple[bool, Optional[str]]:
        """
        Fix merged cells in Excel file using PowerShell.
        
        This method:
        1. Creates temporary PowerShell script
        2. Executes it via subprocess
        3. Returns success status
        
        Args:
            excel_path: Path to Excel file to fix
            
        Returns:
            Tuple of (success: bool, error_message: Optional[str])
        """
        if not ExcelPostProcessor.is_windows():
            return True, None  # Skip on non-Windows
        
        excel_path = str(Path(excel_path).resolve())
        
        try:
            # Create temporary PowerShell script
            temp_script = Path("temp_fix_excel.ps1")
            temp_script.write_text(ExcelPostProcessor.POWERSHELL_SCRIPT, encoding='utf-8')
            
            # Execute PowerShell script
            result = subprocess.run(
                [
                    "powershell.exe",
                    "-ExecutionPolicy", "Bypass",
                    "-File", str(temp_script),
                    "-ExcelPath", excel_path
                ],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            # Cleanup
            if temp_script.exists():
                temp_script.unlink()
            
            # Check result
            if "SUCCESS" in result.stdout:
                return True, None
            else:
                error_msg = result.stderr or result.stdout or "Unknown PowerShell error"
                return False, error_msg
                
        except subprocess.TimeoutExpired:
            return False, "PowerShell script timed out (>60s)"
        except FileNotFoundError:
            return False, "PowerShell not found. Please ensure PowerShell is installed."
        except Exception as e:
            return False, f"Failed to execute PowerShell: {str(e)}"
    
    @staticmethod
    def fix_merged_cells_silent(excel_path: str) -> bool:
        """
        Fix merged cells silently (no error reporting).
        
        Args:
            excel_path: Path to Excel file
            
        Returns:
            True if successful or skipped, False if failed
        """
        success, _ = ExcelPostProcessor.fix_merged_cells(excel_path)
        return success