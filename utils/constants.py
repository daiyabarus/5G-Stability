# ============================================================================
# utils/constants.py
# ============================================================================
"""Column indices and business thresholds for 5G stability reporting."""

from dataclasses import dataclass

class ColumnIndex:
    """Column indices (0-based) for CSV parsing."""
    BEGIN_TIME = 0
    MANAGED_ELEMENT_NAME = 6
    SN_ADDITION_NUM = 21
    ABNORMAL_SN_RELEASE_NUM = 24
    DL_SE_MAPPING_NUM = 35
    DL_UE_THROUGHPUT_NUM = 90
    UL_UE_THROUGHPUT_NUM = 91
    SN_ADDITION_DEN = 22
    ABNORMAL_SN_RELEASE_DEN = 25
    DL_SE_MAPPING_DEN = 36
    DL_UE_THROUGHPUT_DEN = 121
    UL_UE_THROUGHPUT_DEN = 122


@dataclass
class KPIThreshold:
    """KPI thresholds for PASS/FAIL evaluation."""
    SGNB_ADD_SUCCESS_RATE_MIN = 95.0
    SGNB_DROP_RATE_MAX = 5.0
    DL_THROUGHPUT_MIN = 5.0
    UL_THROUGHPUT_MIN = 1.5
    DL_SPECTRUM_EFFICIENCY_MIN = 0.9



TOWERID_PATTERN = r"[A-Z]+-[A-Z]+-[A-Z]+-\d+"

