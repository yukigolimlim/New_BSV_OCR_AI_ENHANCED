"""
lu_analysis_tab.py — entry point shim
=======================================
Now uses the new industry‑based risk model.
Provides dummy placeholders for removed risk‑table constants
so that existing imports do not break.
"""

# Core exports (keep what exists)
from lu_core import (
    GENERAL_CLIENT,
    run_lu_analysis,
    get_product_risk_overrides,
    set_product_risk_overrides,
    _compute_risk_score,
    _cell_str,
    _parse_numeric,
    _fmt_value,
    _numeric_total,
    _RISK_ORDER,
    _SCORE_BANDS,
    _DEFAULT_RISK,
    SECTOR_WHOLESALE,
    SECTOR_AGRICULTURE,
    SECTOR_TRANSPORT,
    SECTOR_REMITTANCE,
    SECTOR_CONSUMER,
    SECTOR_OTHER,
)

# ── These no longer exist in lu_core; provide empty dummies ──
SECTOR_KEYWORDS = {}
SECTOR_KEYWORDS_EXTENDED = {}
EXPENSE_PATTERNS = []
SECTOR_EXPENSE_RISK = {}

# Legacy stub functions (not used, but expected by some old imports)
def _is_summary_sheet(ws):
    return True

def _analyse_sheet(ws, sheet_name):
    return []

def _read_income_from_summary(wb):
    return {}

# ── Attach the UI from lu_ui ────────────────────────────────────
from lu_ui import attach

# Re‑export UI functions that might be expected by other modules
from lu_ui import (
    _build_lu_analysis_panel,
    _build_charts_panel,
    _build_report_panel,
    _build_simulator_panel,
    _build_sim_summary_cards,
    _build_loanbal_panel,
    _loanbal_render,
    _lu_render_results,
    _lu_render_general_view,
    _lu_render_client_view,
    
    _lu_on_client_change,
    _lu_filter_by_search,
    _lu_populate_client_dropdown,
    _lu_switch_view,
    _charts_show_placeholder,
    _charts_render,
    _report_show_placeholder,
    _report_render,
    _report_print,
    _lu_show_export_menu,
    _generate_pdf,
    _generate_excel,
    _sim_show_placeholder,
    _sim_populate,
    _sim_build_expense_row,
    _sim_on_slide,
    _sim_apply_global,
    _sim_reset,
    _sim_refresh,
    _sim_draw_chart,
    _lu_show_placeholder,
    _lu_browse_file,
    _lu_run_analysis,
    _lu_show_error,
    _bind_mousewheel,
    F,
    FF,
)

__all__ = ["attach", "run_lu_analysis", "GENERAL_CLIENT"]