# ============================================================================
# core/data_parser.py
# ============================================================================
"""Parse TOWERID and dates using regex."""

import re
import pandas as pd
from datetime import datetime
from typing import Tuple, List
from utils.constants import TOWERID_PATTERN, ColumnIndex


class DataParser:
    """Extract TOWERID and parse dates from raw data."""

    @staticmethod
    def extract_tower_id(managed_element_name: str) -> str:
        """
        Extract TOWERID using regex pattern: [A-Z]+-[A-Z]+-[A-Z]+-\d+

        Example: "NR_SUM-SU-AKR-0011_DU" -> "SUM-SU-AKR-0011"

        Args:
            managed_element_name: Raw element name string from CSV

        Returns:
            Extracted TOWERID or "UNKNOWN" if pattern not found
        """
        match = re.search(TOWERID_PATTERN, str(managed_element_name))
        return match.group(0) if match else "UNKNOWN"

    @staticmethod
    def parse_date(date_str: str) -> datetime.date:
        """
        Parse date from various formats.

        Supported formats:
        - "1/6/2026 0:00"
        - "1/6/2026"
        - "2026-01-06 00:00:00"
        - "2026-01-06"

        Args:
            date_str: Date string from CSV BEGIN_TIME column

        Returns:
            Parsed date object (defaults to 2026-01-01 if parsing fails)
        """
        try:
            for fmt in ["%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
                try:
                    return datetime.strptime(str(date_str), fmt).date()
                except ValueError:
                    continue

            return pd.to_datetime(date_str).date()
        except:
            return datetime(2026, 1, 1).date()

    @staticmethod
    def add_parsed_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Add TOWERID and DATE columns to DataFrame.

        Args:
            df: Input DataFrame with raw CSV data

        Returns:
            DataFrame with added 'TOWERID' and 'DATE' columns
        """
        df_parsed = df.copy()

        df_parsed["TOWERID"] = df_parsed.iloc[
            :, ColumnIndex.MANAGED_ELEMENT_NAME
        ].apply(DataParser.extract_tower_id)

        df_parsed["DATE"] = df_parsed.iloc[:, ColumnIndex.BEGIN_TIME].apply(
            DataParser.parse_date
        )

        return df_parsed
