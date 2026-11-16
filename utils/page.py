"""
Shared Streamlit page for every agent:
upload -> parse each file -> standardize -> dashboard -> table -> download.

Results are cached per agent, keyed on the exact files uploaded, so
uploading a different set of files always re-processes (no stale data).
"""

import hashlib

import pandas as pd
import streamlit as st

from utils.activity_log import log_event
from utils.dashboard import show_dashboard
from utils.ui import chips, page_title
from utils.common import get_airline
from utils.excel_utils import formula_cells_without_values
from utils.exporter import export_excel
from utils.schema import CANCELLED, is_cancelled, split_issues, standardize


def files_signature(files):
    h = hashlib.md5()
    for f in files:
        h.update(f.name.encode())
        h.update(f.getvalue())
    return h.hexdigest()


BLANK_CHECK = {"AWB_DATE": "date", "DEST": "destination", "CHR_WT": "weight", "TOTAL_FRT": "total"}


def report_blanks(df, f, log, skip_sheet=None):
    """Tell the user exactly which cells came through empty, and why if we can tell."""
    void = is_cancelled(df)
    if void.any():
        log("info", f"{f.name}: {int(void.sum())} cancelled (VOID) AWB(s) left out")
    df = df[~void]
    gaps = []
    for col, label in BLANK_CHECK.items():
        if col in df.columns:
            n = int(pd.Series(df[col]).isna().sum())
            if n:
                gaps.append(f"{n} without {label}")
    if gaps:
        log("warning", f"{f.name}: " + ", ".join(gaps) + " (see DATA_ISSUE column)")

    missing = formula_cells_without_values(f.getvalue(), f.name, skip_sheet)
    if missing:
        log("warning", f"{f.name}: {missing} formula cells have no saved value and read as blank – "
                       "open the file in Excel, press Save, and upload it again")


def process_agent_files(agent_module, files, log):
    """Parse + standardize. Pure (no st.* calls) apart from the log callback."""
    frames = []
    for f in files:
        try:
            df = agent_module.parse(f.getvalue(), f.name, log)
        except Exception as e:  # keep going with the other files
            log("error", f"{f.name}: {e}")
            continue
        if df is None or df.empty:
            log("warning", f"{f.name}: no rows found")
            continue
        df["SOURCE_FILE"] = f.name
        log("info", f"{f.name}: {len(df)} rows")
        report_blanks(df, f, log, getattr(agent_module, "should_skip", None))
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    raw = pd.concat(frames, ignore_index=True, sort=False)
    out = standardize(raw, agent_module.AGENT, agent_module.EXTRA_COLS)
    report_airlines(out, log)
    return out


def report_airlines(df, log):
    """Per file: how many airlines came from the AWB prefix, and which flight cells were not understood."""
    # rows whose bill has no flight at all (PALM: "Not on bill") have nothing to report
    air = df[(df["TRNSPT_MODE"] != "ROAD") & (df["DATA_ISSUE"].fillna("") != CANCELLED)
             & (df["AIRLINE"].fillna("") != "Not on bill")]
    for fname, part in air.groupby("SOURCE_FILE"):
        from_flight = part["FLIGHT_NO"].map(get_airline)
        via_awb = int((from_flight.isna() | (from_flight == "Other")).sum()
                      - (part["AIRLINE"].isna() | (part["AIRLINE"] == "Other")).sum())
        if via_awb:
            log("info", f"{fname}: airline for {via_awb} rows taken from the AWB prefix (flight cell empty/unclear)")
        bad = part[part["AIRLINE"].isna() | (part["AIRLINE"] == "Other")]
        if len(bad):
            samples = bad["FLIGHT_NO"].fillna("(empty)").astype(str).unique()[:6]
            log("warning", f"{fname}: airline not recognised for {len(bad)} rows – flight cells: "
                           + ", ".join(samples) + ". Send these to have them added.")


def show_problems(logs):
    """Errors / warnings stay visible above the results."""
    for lvl, m in logs:
        if lvl == "error":
            st.error(m, icon="⛔")
        elif lvl == "warning":
            st.warning(m, icon="⚠️")


def quality_chips(df, logs, n_files=None, n_excluded=0, n_cancelled=0, n_included=0):
    removed = df.attrs.get("exact_dups_removed", 0)
    dup = int(df["DUP_FLAG"].notna().sum())
    errors = sum(1 for lvl, _ in logs if lvl == "error")
    items = [(f"✓ {len(df):,} rows", "ok")]
    if n_files:
        items.append((f"{n_files} file(s)", "info"))
    if df["AGENT"].nunique() > 1:
        items.append((f"{df['AGENT'].nunique()} agents", "info"))
    if removed:
        items.append((f"{removed} exact repeats removed", ""))
    if dup:
        items.append((f"{dup} repeated AWBs", "warn"))
    if n_cancelled:
        items.append((f"{n_cancelled} cancelled (VOID) left out", ""))
    if n_excluded:
        items.append((f"{n_excluded} rows left out (data issues)", "warn"))
    if n_included:
        items.append((f"{n_included} issue row(s) included by you", "info"))
    if errors:
        items.append((f"{errors} file errors", "bad"))
    chips(items)


ISSUE_COLS = ["AGENT", "INVOICE_NO", "AWB_NO", "AWB_DATE", "OD_PAIR", "CHR_WT", "TOTAL_FRT",
              "DUP_FLAG", "DATA_ISSUE", "SOURCE_FILE"]
DATE_CFG = {"AWB_DATE": st.column_config.DateColumn(format="DD-MMM-YYYY")}


def include_editor(rows, included, key, name):
    """
    Issue rows with an 'Include' tick box. Ticked rows go back into the dashboard,
    data, totals and downloads. `included` (a set of row ids) is kept with this
    upload's results, so a new upload starts with every issue row left out again.
    """
    version = st.session_state.get(f"{key}_inc_v", 0)
    ed_key = f"{key}_inc_{name}_{version}"          # new key after each change -> table redrawn from `included`
    ids = list(rows.index)

    def apply():
        for pos, change in st.session_state[ed_key]["edited_rows"].items():
            if "Include" in change:
                (included.add if change["Include"] else included.discard)(ids[pos])
        st.session_state[f"{key}_inc_v"] = version + 1

    table = rows[ISSUE_COLS].copy()
    table.insert(0, "Include", [i in included for i in ids])
    st.data_editor(
        table, key=ed_key, on_change=apply, width="stretch", hide_index=True,
        disabled=ISSUE_COLS, column_config={**DATE_CFG, "Include": st.column_config.CheckboxColumn(
            "Include", help="Tick to count this row in the dashboard, data, totals and downloads anyway.")},
    )


NO_DATE = "No AWB date"


def month_filter(df, key):
    """'AWB month' picker. Returns (rows for the chosen month, label)."""
    months = df["AWB_DATE"].dt.to_period("M")
    counts = months.value_counts().sort_index(ascending=False)
    labels = {"All months": f"All months ({len(df):,} AWBs)"}
    for p, n in counts.items():
        labels[str(p)] = f"{p.strftime('%b %Y')} ({n:,} AWBs)"
    n_nodate = int(months.isna().sum())
    if n_nodate:
        labels[NO_DATE] = f"{NO_DATE} ({n_nodate:,})"

    choice = st.selectbox(
        "AWB month", list(labels), format_func=labels.get, key=f"{key}_month",
        help="Show only AWBs dated in this month – dashboard, data, issues and the Excel download all follow it.",
    )
    if choice == "All months":
        return df, None
    if choice == NO_DATE:
        return df[months.isna()], NO_DATE
    return df[months.astype(str) == choice], pd.Period(choice).strftime("%b %Y")


def show_results(df_all, logs, key, file_name, make_excel, cache, n_files=None):
    """
    Month filter + status chips + download, then tabs: Dashboard / Data / Issues / Log.
    Rows with a DATA_ISSUE are left out of the dashboard, data and totals; they are
    listed in the Issues tab and on their own sheet in the Excel. The user can tick
    'Include' on any of them in the Issues tab to count it anyway.
    make_excel(clean_rows, excluded_rows) -> bytes
    """
    f_col, _ = st.columns([1, 2])
    with f_col:
        shown, month = month_filter(df_all, key)
    if month:
        hidden = len(df_all) - len(shown)
        st.caption(f"Showing **{month}** only – {hidden:,} AWB(s) from other months hidden.")
    df, flagged = split_issues(shown)
    is_void = flagged["DATA_ISSUE"] == CANCELLED
    all_voided, all_issues = flagged[is_void], flagged[~is_void]

    # issue rows the user ticked 'Include' in the Issues tab go back into the totals
    included = cache.setdefault("included", set())
    chosen = flagged.index.isin(included)
    if chosen.any():
        attrs = df.attrs
        df = pd.concat([df, flagged[chosen]]).sort_index()
        df.attrs = attrs
    excluded = flagged[~chosen]
    voided, issues = excluded[excluded["DATA_ISSUE"] == CANCELLED], excluded[excluded["DATA_ISSUE"] != CANCELLED]

    # Excel follows the month filter and the included rows; each version is built once and kept
    tag = (month or "all") + "|" + ",".join(map(str, sorted(flagged.index[chosen])))
    xl = cache.setdefault("xlsx_by_month", {})
    if tag not in xl:
        with st.spinner("Preparing Excel…"):
            xl[tag] = make_excel(df, excluded)
    out_name = file_name if not month else file_name.replace(".xlsx", f"_{month.replace(' ', '')}.xlsx")

    top_l, top_r = st.columns([3, 1], vertical_alignment="center")
    with top_l:
        quality_chips(df, logs, n_files, n_excluded=len(issues), n_cancelled=len(voided), n_included=int(chosen.sum()))
    with top_r:
        st.download_button(
            "⬇  Download Excel",
            data=xl[tag],
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key}_download",
            type="primary",
            width="stretch",
            on_click=log_event,
            args=("DOWNLOAD", out_name),
        )
    show_problems(logs)
    # always draw this slot (even when empty) so the tabs below keep their place – otherwise ticking
    # 'Include' on the last left-out row removes the note and the page jumps back to the Dashboard tab
    note = st.empty()
    if len(excluded):
        parts = []
        if len(voided):
            parts.append(f"{len(voided):,} cancelled (VOID) AWB(s)")
        if len(issues):
            parts.append(f"{len(issues):,} row(s) with data issues")
        note.info(" and ".join(parts) + " were left out of the dashboard, data and totals – see the **Issues** tab "
                "(also on the *Excluded_Issues* sheet in the Excel). Tick **Include** there to count a row anyway.",
                icon="ℹ️")

    dups = df[df["DUP_FLAG"].notna()]
    infos = [m for lvl, m in logs if lvl == "info"]
    labels = ["📊  Dashboard", "📋  Data", f"⚠️  Issues ({len(flagged) + len(dups)})", "🧾  Processing log"]
    try:   # remembered tab, so ticking 'Include' keeps the Issues tab open (newer Streamlit only)
        t_dash, t_data, t_issue, t_log = st.tabs(labels, key=f"{key}_tabs", on_change="rerun")
    except TypeError:
        t_dash, t_data, t_issue, t_log = st.tabs(labels)

    with t_dash:
        if df.empty:
            st.warning("Every row has a data issue, so there is nothing to chart – see the Issues tab.")
        else:
            show_dashboard(df, key=key, pdf_name=out_name.replace(".xlsx", "_Dashboard.pdf"), month=month)

    with t_data:
        q = st.text_input("Search", placeholder="AWB no, invoice, lane…", key=f"{key}_q",
                          label_visibility="collapsed")
        view = df
        if q:
            mask = pd.Series(False, index=df.index)
            for col in ["AWB_NO", "INVOICE_NO", "OD_PAIR", "FLIGHT_NO", "SOURCE_FILE"]:
                mask |= df[col].astype(str).str.contains(q, case=False, na=False, regex=False)
            view = df[mask]
        st.caption(f"{len(view):,} of {len(df):,} rows")
        st.dataframe(view, width="stretch", height=520, column_config=DATE_CFG)

    with t_issue:
        if flagged.empty and dups.empty:
            st.success("No repeated AWBs or data issues found.", icon="✅")
        if len(all_issues):
            st.markdown(f"**Data issues ({len(all_issues):,}) – {len(issues):,} left out**")
            st.caption("Something missing or not matching in the bill, so these rows are left out of the dashboard, "
                       "data and totals. Tick **Include** to count a row anyway (it keeps its DATA_ISSUE note), "
                       "or fix the bill and upload it again.")
            include_editor(all_issues, included, key, "issues")
        if len(all_voided):
            st.markdown(f"**Cancelled – VOID ({len(all_voided):,}) – {len(voided):,} left out**")
            st.caption("The bill marks these AWBs as VOID (cancelled), so they are not counted. "
                       "Tick **Include** to count one anyway.")
            include_editor(all_voided, included, key, "void")
        if len(dups):
            st.markdown(f"**Repeated AWBs ({len(dups):,})**")
            st.caption("Same AWB on more than one row. These rows are kept in the totals.")
            st.dataframe(dups[ISSUE_COLS], width="stretch", hide_index=True, column_config=DATE_CFG)

    with t_log:
        if not infos:
            st.caption("Nothing to show.")
        for m in infos:
            st.markdown(f"- {m}")


def run_agent_page(agent_module):
    key = f"agent_{agent_module.KEY}"
    types = ", ".join(f".{t}" for t in agent_module.FILE_TYPES)
    page_title(f"{agent_module.LABEL}", getattr(agent_module, "DESCRIPTION", "") or f"Upload {types} bills")

    with st.container(border=True):
        files = st.file_uploader(
            f"Upload {agent_module.LABEL} bills ({types}) – you can select several files",
            type=agent_module.FILE_TYPES,
            accept_multiple_files=True,
            key=f"{key}_upload",
        )

    if not files:
        st.session_state.pop(key, None)
        st.caption("Results appear here once files are uploaded.")
        return

    sig = files_signature(files)
    cached = st.session_state.get(key)

    if not cached or cached["sig"] != sig:
        logs = []
        with st.spinner(f"Reading {len(files)} file(s)…"):
            df = process_agent_files(agent_module, files, lambda lvl, m: logs.append((lvl, m)))
        cached = {"sig": sig, "df": df, "logs": logs}
        st.session_state[key] = cached
        log_event("PROCESS", f"{agent_module.LABEL}: {len(files)} file(s), {len(df)} rows – " + ", ".join(f.name for f in files))

    df = cached["df"]
    if df.empty:
        show_problems(cached["logs"])
        st.warning("No valid rows found in these files.")
        return

    sheet = agent_module.KEY.upper()
    show_results(
        df, cached["logs"], key, agent_module.OUTPUT_FILE,
        lambda d, bad: export_excel(d, sheet_name=sheet, excluded=bad).getvalue(), cached, n_files=len(files),
    )
