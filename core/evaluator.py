# ============================================================================
# core/evaluator.py
# ============================================================================
"""Evaluate PASS/FAIL status based on KPI thresholds."""

from typing import Tuple
from utils.constants import KPIThreshold
from models.data_models import KPIValues, KPIStatus


class Evaluator:
    """Evaluate KPIs against business thresholds."""
    
    @staticmethod
    def evaluate_kpi(value: float, threshold: float, 
                     compare: str = ">=") -> str:
        """
        Evaluate single KPI against threshold.
        
        Args:
            value: KPI value to evaluate
            threshold: Threshold value for comparison
            compare: Comparison operator (">=" or "<=")
            
        Returns:
            "PASS", "FAIL", or "N/A" (if value is None)
        """
        if value is None:
            return "N/A"
        
        if compare == ">=":
            return "PASS" if value >= threshold else "FAIL"
        else:
            return "PASS" if value <= threshold else "FAIL"
    
    @staticmethod
    def evaluate_all(kpi_values: KPIValues) -> Tuple[KPIStatus, str]:
        """
        Evaluate all KPIs and determine overall status.
        
        PASS/FAIL Thresholds:
        - SgNB Add Success Rate >= 95% → PASS
        - SgNB Drop Rate <= 5% → PASS
        - DL User Throughput >= 5 Mbps → PASS
        - UL User Throughput >= 1.5 Mbps → PASS
        - DL Spectrum Efficiency >= 0.9 → PASS
        
        Overall Status:
        - FAIL if ANY KPI is FAIL
        - PASS only if ALL KPIs are PASS
        
        Args:
            kpi_values: Calculated KPI values
            
        Returns:
            Tuple of (KPIStatus object, overall_status string)
        """
        thresh = KPIThreshold()

        status = KPIStatus(
            sgnb_add_success_rate=Evaluator.evaluate_kpi(
                kpi_values.sgnb_add_success_rate, 
                thresh.SGNB_ADD_SUCCESS_RATE_MIN, 
                ">="
            ),
            sgnb_drop_rate=Evaluator.evaluate_kpi(
                kpi_values.sgnb_drop_rate,
                thresh.SGNB_DROP_RATE_MAX,
                "<="
            ),
            dl_user_throughput=Evaluator.evaluate_kpi(
                kpi_values.dl_user_throughput,
                thresh.DL_THROUGHPUT_MIN,
                ">="
            ),
            ul_user_throughput=Evaluator.evaluate_kpi(
                kpi_values.ul_user_throughput,
                thresh.UL_THROUGHPUT_MIN,
                ">="
            ),
            dl_spectrum_efficiency=Evaluator.evaluate_kpi(
                kpi_values.dl_spectrum_efficiency,
                thresh.DL_SPECTRUM_EFFICIENCY_MIN,
                ">="
            )
        )

        statuses = [
            status.sgnb_add_success_rate,
            status.sgnb_drop_rate,
            status.dl_user_throughput,
            status.ul_user_throughput,
            status.dl_spectrum_efficiency
        ]
        
        overall = "FAIL" if "FAIL" in statuses else "PASS"
        
        return status, overall
