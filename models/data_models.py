# ============================================================================
# models/data_models.py
# ============================================================================
"""Data models for type-safe processing."""

from dataclasses import dataclass
from datetime import date
from typing import Optional


@dataclass
class KPIValues:
    """Calculated KPI values for a tower on a specific date."""
    sgnb_add_success_rate: Optional[float] = None
    sgnb_drop_rate: Optional[float] = None
    dl_user_throughput: Optional[float] = None
    ul_user_throughput: Optional[float] = None
    dl_spectrum_efficiency: Optional[float] = None


@dataclass
class KPIStatus:
    """PASS/FAIL status for each KPI."""
    sgnb_add_success_rate: str = "N/A"
    sgnb_drop_rate: str = "N/A"
    dl_user_throughput: str = "N/A"
    ul_user_throughput: str = "N/A"
    dl_spectrum_efficiency: str = "N/A"


@dataclass
class TowerReport:
    """Complete report for a single tower on a specific date."""
    tower_id: str
    report_date: date
    kpi_values: KPIValues
    kpi_status: KPIStatus
    overall_status: str