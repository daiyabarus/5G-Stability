# ============================================================================
# core/data_loader.py
# ============================================================================
"""Load and filter CSV files from ZIP archives."""

import zipfile
import pandas as pd
from typing import List


class DataLoader:
    """Handles ZIP extraction and CSV loading."""

    @staticmethod
    def load_from_zip(zip_path: str) -> pd.DataFrame:
        """
        Load all CSVs from ZIP except *_KPI(Counter).csv files.

        Args:
            zip_path: Path to the ZIP file containing CSV files

        Returns:
            Combined DataFrame from all valid CSV files

        Raises:
            ValueError: If no valid CSV files found in ZIP
        """
        dfs = []

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            csv_files = [
                f
                for f in zip_ref.namelist()
                if f.endswith(".csv") and not f.endswith("_KPI(Counter).csv")
            ]

            if not csv_files:
                raise ValueError("No valid CSV files found in ZIP archive")

            for csv_file in csv_files:
                with zip_ref.open(csv_file) as f:
                    df = pd.read_csv(f, header=0, low_memory=False)
                    dfs.append(df)

        return pd.concat(dfs, ignore_index=True)
