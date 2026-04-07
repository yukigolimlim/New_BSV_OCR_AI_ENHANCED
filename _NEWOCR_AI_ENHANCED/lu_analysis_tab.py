"""
lu_analysis_tab.py — entry point shim
======================================
Re-exports the public surface of lu_core and lu_ui so any existing
code that does:

    import lu_analysis_tab
    lu_analysis_tab.attach(DocExtractorApp)

continues to work without modification.
"""

from lu_core import (
    SECTOR_KEYWORDS, SECTOR_KEYWORDS_EXTENDED, EXPENSE_PATTERNS,
    SECTOR_EXPENSE_RISK, GENERAL_CLIENT, run_lu_analysis,
    _compute_risk_score, _cell_str, _parse_numeric, _fmt_value,
    _numeric_total, _is_summary_sheet, _analyse_sheet,
    _read_income_from_summary, _RISK_ORDER, _SCORE_BANDS, _DEFAULT_RISK,
    SECTOR_WHOLESALE, SECTOR_AGRICULTURE, SECTOR_TRANSPORT,
    SECTOR_REMITTANCE, SECTOR_CONSUMER, SECTOR_OTHER,
)

from lu_ui import (
    attach,
    _build_lu_analysis_panel, _build_charts_panel,
    _build_report_panel, _build_simulator_panel, _build_sim_summary_cards,
    _build_loanbal_panel, _loanbal_render, 
    _lu_render_results, _lu_render_general_view, _lu_render_client_view,
    _lu_render_client_card, _lu_render_sector_card,
    _lu_on_client_change, _lu_filter_by_search, _lu_populate_client_dropdown,
    _lu_switch_view, _charts_show_placeholder, _charts_render,
    _report_show_placeholder, _report_render, _report_print,
    _lu_show_export_menu, _generate_pdf, _generate_excel,
    _sim_show_placeholder, _sim_populate, _sim_build_expense_row,
    _sim_on_slide, _sim_apply_global, _sim_reset, _sim_refresh, _sim_draw_chart,
    _lu_show_placeholder, _lu_browse_file, _lu_run_analysis, _lu_show_error,
    _bind_mousewheel, F, FF,
)

__all__ = ["attach", "run_lu_analysis", "GENERAL_CLIENT"]