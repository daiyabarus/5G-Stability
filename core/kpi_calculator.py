# ============================================================================
# core/kpi_calculator.py
# ============================================================================
"""Calculate KPIs from numerator and denominator columns."""

import pandas as pd
from typing import Optional
from utils.constants import ColumnIndex
from models.data_models import KPIValues


class KPICalculator:
    """Calculate 5G KPIs with safe division (zero-division handling)."""

    @staticmethod
    def safe_divide(
        num: Optional[float], den: Optional[float], multiply: float = 1.0
    ) -> Optional[float]:
        """
        Safely divide with zero-check and NaN handling.

        Args:
            num: Numerator value
            den: Denominator value
            multiply: Multiplier for result (e.g., 100 for percentages)

        Returns:
            Calculated result or None if division is impossible
        """
        if pd.isna(num) or pd.isna(den) or den == 0:
            return None
        return (num / den) * multiply

    @staticmethod
    def calculate_kpis(group: pd.DataFrame) -> KPIValues:
        """
        Calculate aggregated KPIs for a group (TOWERID + DATE).

        Process:
        1. Sum all numerators for the group
        2. Sum all denominators for the group
        3. Calculate KPI = SUM(NUM) / SUM(DEN)

        KPI Formulas:
        - SgNB Add Success Rate (%) = SN_ADDITION_NUM / SN_ADDITION_DEN * 100
        - SgNB Drop Rate (%) = ABNORMAL_SN_RELEASE_NUM / ABNORMAL_SN_RELEASE_DEN * 100
        - DL User Throughput (Mbps) = DL_UE_THROUGHPUT_NUM / DL_UE_THROUGHPUT_DEN
        - UL User Throughput (Mbps) = UL_UE_THROUGHPUT_NUM / UL_UE_THROUGHPUT_DEN
        - DL Spectrum Efficiency = DL_SE_MAPPING_NUM / DL_SE_MAPPING_DEN

        Args:
            group: DataFrame group for a specific tower and date

        Returns:
            KPIValues object with calculated metrics
        """
        sn_add_num = group.iloc[:, ColumnIndex.SN_ADDITION_NUM].sum()
        sn_add_den = group.iloc[:, ColumnIndex.SN_ADDITION_DEN].sum()

        abn_rel_num = group.iloc[:, ColumnIndex.ABNORMAL_SN_RELEASE_NUM].sum()
        abn_rel_den = group.iloc[:, ColumnIndex.ABNORMAL_SN_RELEASE_DEN].sum()

        dl_tp_num = group.iloc[:, ColumnIndex.DL_UE_THROUGHPUT_NUM].sum()
        dl_tp_den = group.iloc[:, ColumnIndex.DL_UE_THROUGHPUT_DEN].sum()

        ul_tp_num = group.iloc[:, ColumnIndex.UL_UE_THROUGHPUT_NUM].sum()
        ul_tp_den = group.iloc[:, ColumnIndex.UL_UE_THROUGHPUT_DEN].sum()

        dl_se_num = group.iloc[:, ColumnIndex.DL_SE_MAPPING_NUM].sum()
        dl_se_den = group.iloc[:, ColumnIndex.DL_SE_MAPPING_DEN].sum()

        return KPIValues(
            sgnb_add_success_rate=KPICalculator.safe_divide(
                sn_add_num, sn_add_den, 100
            ),
            sgnb_drop_rate=KPICalculator.safe_divide(abn_rel_num, abn_rel_den, 100),
            dl_user_throughput=KPICalculator.safe_divide(dl_tp_num, dl_tp_den),
            ul_user_throughput=KPICalculator.safe_divide(ul_tp_num, ul_tp_den),
            dl_spectrum_efficiency=KPICalculator.safe_divide(dl_se_num, dl_se_den),
        )
