# ============================================================================
# gui/main_window.py
# ============================================================================
"""PyQt6 GUI for the 5G Stability Reporter."""

import sys
from pathlib import Path
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QTextEdit, QFileDialog, QMessageBox, QLabel
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont
import traceback

from core.data_loader import DataLoader
from core.data_cleaner import DataCleaner
from core.data_parser import DataParser
from core.kpi_calculator import KPICalculator
from core.evaluator import Evaluator
from core.excel_generator import ExcelGenerator
from utils.constants import ColumnIndex
from models.data_models import TowerReport


class ProcessingThread(QThread):
    """Background thread for data processing to keep UI responsive."""
    
    # Signals for communication with main thread
    progress = pyqtSignal(str)   # Progress messages
    finished = pyqtSignal(dict)  # Completion with results
    error = pyqtSignal(str)      # Error messages
    
    def __init__(self, zip_path: str, output_path: str):
        """
        Initialize processing thread.
        
        Args:
            zip_path: Path to input ZIP file
            output_path: Path for output Excel file
        """
        super().__init__()
        self.zip_path = zip_path
        self.output_path = output_path
    
    def run(self):
        """Execute the complete processing pipeline."""
        try:
            # Step 1: Load data from ZIP
            self.progress.emit("📂 Loading CSV files from ZIP...")
            df = DataLoader.load_from_zip(self.zip_path)
            self.progress.emit(f"✓ Loaded {len(df)} records from CSV files")
            
            # Step 2: Clean numeric columns
            self.progress.emit("🧹 Cleaning numeric columns (removing %, commas)...")
            numeric_cols = [
                ColumnIndex.SN_ADDITION_NUM, ColumnIndex.SN_ADDITION_DEN,
                ColumnIndex.ABNORMAL_SN_RELEASE_NUM, ColumnIndex.ABNORMAL_SN_RELEASE_DEN,
                ColumnIndex.DL_UE_THROUGHPUT_NUM, ColumnIndex.DL_UE_THROUGHPUT_DEN,
                ColumnIndex.UL_UE_THROUGHPUT_NUM, ColumnIndex.UL_UE_THROUGHPUT_DEN,
                ColumnIndex.DL_SE_MAPPING_NUM, ColumnIndex.DL_SE_MAPPING_DEN
            ]
            df_clean = DataCleaner.clean_dataframe(df, numeric_cols)
            self.progress.emit("✓ Data cleaning completed")
            
            # Step 3: Parse TOWERID and dates using regex
            self.progress.emit("🔍 Parsing TOWERID (regex) and dates...")
            df_parsed = DataParser.add_parsed_columns(df_clean)
            self.progress.emit("✓ TOWERID and DATE columns added")
            
            # Step 4: Group by (TOWERID, DATE) and calculate KPIs
            self.progress.emit("📊 Calculating KPIs (grouping by TOWERID + DATE)...")
            grouped = df_parsed.groupby(['TOWERID', 'DATE'])
            reports = []
            
            for (tower_id, dt), group in grouped:
                # Calculate KPIs for this tower-date combination
                kpi_values = KPICalculator.calculate_kpis(group)
                
                # Evaluate PASS/FAIL status
                kpi_status, overall = Evaluator.evaluate_all(kpi_values)
                
                # Create report object
                reports.append(TowerReport(tower_id, dt, kpi_values, kpi_status, overall))
            
            self.progress.emit(f"✓ Processed {len(reports)} tower-date combinations")
            
            # Step 5: Generate Excel report with styling
            self.progress.emit("📝 Generating Excel report with color coding...")
            ExcelGenerator.create_report(reports, self.output_path)
            self.progress.emit("✓ Excel report generated successfully")
            
            # Calculate statistics
            pass_count = sum(1 for r in reports if r.overall_status == "PASS")
            fail_count = sum(1 for r in reports if r.overall_status == "FAIL")
            
            # Emit completion signal with results
            self.finished.emit({
                'total': len(reports),
                'pass': pass_count,
                'fail': fail_count,
                'output': self.output_path
            })
            
        except Exception as e:
            # Emit error signal with traceback
            self.error.emit(f"Error: {str(e)}\n\n{traceback.format_exc()}")


class MainWindow(QMainWindow):
    """Main application window with modern UI."""
    
    def __init__(self):
        """Initialize the main window."""
        super().__init__()
        self.zip_path = None
        self.processing_thread = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize user interface components."""
        self.setWindowTitle("5G Stability Report Generator")
        self.setGeometry(100, 100, 700, 500)
        
        # Create central widget and layout
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title label
        title = QLabel("5G Performance Stability Report Generator")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # File selection status label
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: #666; padding: 10px;")
        layout.addWidget(self.file_label)
        
        # Browse ZIP button
        self.browse_btn = QPushButton("Browse ZIP File")
        self.browse_btn.setMinimumHeight(40)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.browse_btn.clicked.connect(self.browse_file)
        layout.addWidget(self.browse_btn)
        
        # Start processing button
        self.start_btn = QPushButton("START PROCESSING")
        self.start_btn.setMinimumHeight(50)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 16px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover:enabled {
                background-color: #0b7dda;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)
        
        # Log area label
        layout.addWidget(QLabel("Processing Log:"))
        
        # Log text area
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 10px;
                font-family: 'Courier New';
            }
        """)
        layout.addWidget(self.log_area)
    
    def browse_file(self):
        """Open file dialog to select ZIP file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select 5G Performance ZIP File",
            "",
            "ZIP Files (*.zip)"
        )
        
        if file_path:
            self.zip_path = file_path
            self.file_label.setText(f"Selected: {Path(file_path).name}")
            self.file_label.setStyleSheet("color: #4CAF50; padding: 10px; font-weight: bold;")
            self.start_btn.setEnabled(True)
            self.log("📁 File selected: " + file_path)
    
    def start_processing(self):
        """Start processing in background thread."""
        if not self.zip_path:
            return
        
        # Generate default filename with datetime
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"5G_Stability_Report_{current_time}.xlsx"
        
        # Get output file path from user
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report As",
            default_filename,  # Pre-populate with datetime filename
            "Excel Files (*.xlsx)"
        )
        
        if not output_path:
            return
        
        # Ensure file has .xlsx extension
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'
        
        # Disable buttons during processing
        self.browse_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.log_area.clear()
        
        # Create and start background thread
        self.processing_thread = ProcessingThread(self.zip_path, output_path)
        self.processing_thread.progress.connect(self.log)
        self.processing_thread.finished.connect(self.on_finished)
        self.processing_thread.error.connect(self.on_error)
        self.processing_thread.start()
    
    def log(self, message: str):
        """Append message to log area."""
        self.log_area.append(message)
    
    def on_finished(self, results: dict):
        """Handle successful completion."""
        # Re-enable buttons
        self.browse_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        
        # Show success message
        msg = (
            f"✅ Processing Complete!\n\n"
            f"Total Tower-Date Combinations: {results['total']}\n"
            f"PASS: {results['pass']}\n"
            f"FAIL: {results['fail']}\n\n"
            f"Report saved to:\n{results['output']}"
        )
        
        QMessageBox.information(self, "Success", msg)
        self.log("\n" + "="*50)
        self.log("🎉 PROCESSING COMPLETE!")
        self.log("="*50)
    
    def on_error(self, error_msg: str):
        """Handle processing error."""
        # Re-enable buttons
        self.browse_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        
        # Show error dialog
        QMessageBox.critical(self, "Processing Error", error_msg)
        self.log("\n" + "="*50)
        self.log("❌ ERROR OCCURRED")
        self.log("="*50)
        self.log(error_msg)