# ============================================================================
# core/data_cleaner.py
# ============================================================================
"""Clean numeric fields by removing formatting characters."""

import pandas as pd
from typing import List


class DataCleaner:
    """Clean numeric columns by removing %, commas, and converting to float."""

    @staticmethod
    def clean_dataframe(df: pd.DataFrame, numeric_columns: List[int]) -> pd.DataFrame:
        """
        Clean numeric columns in the DataFrame.

        Removes:
        - Percent symbols (%)
        - Thousand separators (,)
        - Empty strings

        Converts to numeric type with NaN for invalid values.

        Args:
            df: Input DataFrame
            numeric_columns: List of column indices to clean

        Returns:
            Cleaned DataFrame with numeric columns converted
        """
        df_clean = df.copy()

        for col_idx in numeric_columns:
            if col_idx < len(df_clean.columns):
                col = df_clean.iloc[:, col_idx]

                df_clean.iloc[:, col_idx] = (
                    col.astype(str)
                    .str.replace("%", "", regex=False)
                    .str.replace(",", "", regex=False)
                    .replace("", None)
                )

                df_clean.iloc[:, col_idx] = pd.to_numeric(
                    df_clean.iloc[:, col_idx], errors="coerce"
                )

        return df_clean
