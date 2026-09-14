"""AKS Migration App — Streamlit front end.

This module is presentation only. Every piece of migration behaviour lives in
the `migration` package and is reached through two calls:

    analysis = analyse_migration(source, mapping_path=..., template_path=...)
    result   = run_migration(analysis, highlight=..., replace_template_data=...)

Nothing here reads a workbook, talks to Excel, or decides what a value becomes.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from migration import (  # noqa: E402
    LOGS_DIR,
    OUTPUTS_DIR,
    STATUS_OK,
    STATUS_REVIEW,
    STATUS_SKIPPED,
    STATUS_WARNING,
    __version__,
    analyse_migration,
    check_environment,
    cleanup_staged_uploads,
    default_mapping,
    default_template,
    describe_configuration,
    inspect_source_file,
    planned_output_name,
    previous_outputs,
    run_migration,
    stage_configuration_file,
    stage_source,
    version_label,
)
from migration.paths import CONFIG_DIR, TEMPLATES_DIR, ensure_directories  # noqa: E402

APP_TITLE = "AKS Migration App"
APP_SUBTITLE = "Convert legacy AKS lists to the AKS V2 format"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

STYLE = """
<style>
  .block-container {padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1500px;}
  h1 {margin-bottom: 0.1rem;}
  .aks-subtitle {color: #5b6b7c; font-size: 1.05rem; margin-bottom: 0.4rem;}
  .aks-step {color: #0b4f8a; font-weight: 700; font-size: 0.82rem;
             letter-spacing: 0.09em; text-transform: uppercase; margin-bottom: 0.2rem;}
  .aks-pill {display: inline-block; padding: 0.12rem 0.55rem; border-radius: 999px;
             font-size: 0.78rem; font-weight: 600; margin-right: 0.35rem;}
  .aks-ok   {background: #dff3e3; color: #17632a;}
  .aks-rev  {background: #fde2e4; color: #97202c;}
  .aks-warn {background: #fdf0cd; color: #8a6100;}
  .aks-muted{background: #eef1f4; color: #5b6b7c;}
  .aks-note {color: #5b6b7c; font-size: 0.86rem;}
  div[data-testid="stMetricValue"] {font-size: 1.7rem;}
</style>
"""
st.markdown(STYLE, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

DEFAULTS = {
    "source_path": None,
    "source_name": None,
    "source_signature": None,
    "inspection": None,
    "analysis": None,
    "result": None,
    "mapping_path": None,
    "template_path": None,
}
for key, value in DEFAULTS.items():
    st.session_state.setdefault(key, value)


def reset_downstream(keep_analysis: bool = False) -> None:
    """A changed input invalidates everything computed from it."""
    if not keep_analysis:
        st.session_state.analysis = None
    st.session_state.result = None


def open_in_explorer(path: Path) -> None:
    """The app runs on the user's own PC, so opening a folder is safe and useful."""
    try:
        os.startfile(str(Path(path)))  # noqa: S606 - local Windows desktop use
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Could not open the folder. It is here: {path}  ({exc})")


def pill(text: str, kind: str = "muted") -> str:
    return f'<span class="aks-pill aks-{kind}">{text}</span>'


def status_frame(rows: list[dict]) -> pd.DataFrame:
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    return frame.astype({column: "string" for column in frame.columns if column != "Source Row"})


# ---------------------------------------------------------------------------
# Header and sidebar
# ---------------------------------------------------------------------------

ensure_directories()

st.title(APP_TITLE)
st.markdown(f'<div class="aks-subtitle">{APP_SUBTITLE}</div>', unsafe_allow_html=True)
st.divider()


@st.cache_data(show_spinner=False, ttl=600)
def cached_environment():
    report = check_environment()
    return report.ok, [(check.name, check.ok, check.detail) for check in report.checks]


with st.sidebar:
    st.subheader("System check")
    env_ok, env_checks = cached_environment()
    for name, ok, detail in env_checks:
        if ok:
            st.markdown(f"✅ **{name}**  \n<span class='aks-note'>{detail}</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"❌ **{name}**  \n<span class='aks-note'>{detail}</span>", unsafe_allow_html=True)
    if not env_ok:
        st.error(
            "The analysis still works, but creating a migrated workbook needs Microsoft Excel "
            "for Windows on this PC."
        )
    if st.button("Re-check", use_container_width=True):
        cached_environment.clear()
        st.rerun()

    st.divider()
    st.subheader("Folders")
    if st.button("Open results folder", use_container_width=True):
        open_in_explorer(OUTPUTS_DIR)
    if st.button("Open log folder", use_container_width=True):
        open_in_explorer(LOGS_DIR)

    st.divider()
    st.caption(f"AKS Migration App v{__version__}")
    st.caption("Deterministic migration. Unknown values are flagged, never guessed.")


# ---------------------------------------------------------------------------
# Step 1 — select the old AKS file
# ---------------------------------------------------------------------------

st.markdown('<div class="aks-step">Step 1</div>', unsafe_allow_html=True)
st.subheader("Select the old AKS file")

upload = st.file_uploader(
    "Legacy AKS workbook (.xlsb)",
    type=["xlsb"],
    help="Your original file is never changed. The app works on its own copy.",
)

if upload is not None:
    # Streamlit gives every upload its own file_id. Keying on name and size
    # alone would silently keep the previously staged copy when a corrected
    # workbook is re-uploaded under the same name with the same byte count.
    signature = (upload.name, upload.size, getattr(upload, "file_id", None))
    if st.session_state.source_signature != signature:
        with st.spinner("Reading the workbook…"):
            staged = stage_source(upload.getvalue(), upload.name)
            st.session_state.source_path = staged
            st.session_state.source_name = upload.name
            st.session_state.source_signature = signature
            st.session_state.inspection = inspect_source_file(staged)
            reset_downstream()
            cleanup_staged_uploads(keep_run_id=staged.parent.name)
elif st.session_state.source_signature is not None:
    for key, value in DEFAULTS.items():
        st.session_state[key] = value

inspection = st.session_state.inspection

if inspection is None:
    st.info("Choose a legacy AKS `.xlsb` file to begin.")
elif not inspection.ok:
    st.error("**This file cannot be used as a legacy AKS list.**")
    for problem in inspection.problems:
        st.markdown(f"- {problem}")
    st.caption(
        "A legacy AKS list is an Excel binary workbook (.xlsb) containing an 'Übersicht' "
        "worksheet with the equipment table."
    )
else:
    with st.container(border=True):
        columns = st.columns([3, 2, 2, 2])
        columns[0].markdown(f"**File**  \n{st.session_state.source_name}")
        columns[1].markdown(f"**Worksheet**  \n{inspection.sheet}")
        columns[2].metric("Equipment rows", f"{inspection.data_rows:,}")
        columns[3].markdown(
            f"**Validation**  \n{pill('Valid legacy AKS list', 'ok')}", unsafe_allow_html=True
        )
        details = [
            f"Data read from rows {inspection.first_row}–{inspection.last_row}.",
            f"Column headings found on row(s) {', '.join(str(row) for row in inspection.header_rows_found)}.",
        ]
        details.extend(inspection.notes)
        st.markdown(
            "<span class='aks-note'>" + " ".join(details) + "</span>", unsafe_allow_html=True
        )


# ---------------------------------------------------------------------------
# Step 2 — migration configuration
# ---------------------------------------------------------------------------

st.markdown('<div class="aks-step">Step 2</div>', unsafe_allow_html=True)
st.subheader("Migration configuration")


def workbook_choices(directory: Path, pattern: str, shipped: Path) -> list[Path]:
    found = sorted(path for path in directory.glob(pattern) if path.is_file())
    ordered = [shipped] if shipped.exists() else []
    ordered.extend(path for path in found if path.resolve() != shipped.resolve())
    return ordered


shipped_mapping = default_mapping()
shipped_template = default_template()

with st.container(border=True):
    left, right = st.columns(2)
    left.markdown(f"**Migration mapping**  \n{version_label(st.session_state.mapping_path or shipped_mapping)}")
    right.markdown(f"**AKS V2 template**  \n{version_label(st.session_state.template_path or shipped_template)}")
    if st.session_state.mapping_path or st.session_state.template_path:
        st.markdown(
            f"{pill('Custom configuration in use', 'warn')}"
            "<span class='aks-note'>An advanced setting replaces a file shipped with the app.</span>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"{pill('Standard configuration', 'ok')}"
            "<span class='aks-note'>The files shipped with the application are used.</span>",
            unsafe_allow_html=True,
        )

with st.expander("Advanced settings — use a different mapping or template"):
    st.caption(
        "Only change these when the engineering team has released a new mapping or a new AKS V2 "
        "template. Everything else in the migration stays the same."
    )
    advanced_left, advanced_right = st.columns(2)

    with advanced_left:
        mapping_options = workbook_choices(CONFIG_DIR, "*.xlsx", shipped_mapping)
        mapping_labels = [f"{path.name} (shipped)" if index == 0 else path.name
                          for index, path in enumerate(mapping_options)]
        chosen = st.selectbox(
            "Migration mapping file",
            options=list(range(len(mapping_options))),
            format_func=lambda index: mapping_labels[index],
        ) if mapping_options else None
        mapping_upload = st.file_uploader("…or upload another mapping (.xlsx)", type=["xlsx"], key="mapping_upload")

    with advanced_right:
        template_options = workbook_choices(TEMPLATES_DIR, "*.xlsb", shipped_template)
        template_labels = [f"{path.name} (shipped)" if index == 0 else path.name
                           for index, path in enumerate(template_options)]
        chosen_template = st.selectbox(
            "AKS V2 template file",
            options=list(range(len(template_options))),
            format_func=lambda index: template_labels[index],
        ) if template_options else None
        template_upload = st.file_uploader("…or upload another template (.xlsb)", type=["xlsb"], key="template_upload")

    new_mapping = None
    if mapping_upload is not None:
        new_mapping = stage_configuration_file(mapping_upload.getvalue(), mapping_upload.name)
    elif chosen is not None and mapping_options[chosen].resolve() != shipped_mapping.resolve():
        new_mapping = mapping_options[chosen]

    new_template = None
    if template_upload is not None:
        new_template = stage_configuration_file(template_upload.getvalue(), template_upload.name)
    elif chosen_template is not None and template_options[chosen_template].resolve() != shipped_template.resolve():
        new_template = template_options[chosen_template]

    if (new_mapping, new_template) != (st.session_state.mapping_path, st.session_state.template_path):
        st.session_state.mapping_path = new_mapping
        st.session_state.template_path = new_template
        reset_downstream()
        st.rerun()

configuration = describe_configuration(st.session_state.mapping_path, st.session_state.template_path)
for problem in configuration["mapping_problems"] + configuration["template_problems"]:
    st.error(problem)

configuration_ok = not (configuration["mapping_problems"] or configuration["template_problems"])


# ---------------------------------------------------------------------------
# Step 3 — analyse (dry run)
# ---------------------------------------------------------------------------

st.markdown('<div class="aks-step">Step 3</div>', unsafe_allow_html=True)
st.subheader("Analyse migration")
st.caption(
    "A dry run. Every legacy value is compared with the migration mapping and the AKS V2 "
    "value lists. Nothing is written."
)

ready_to_analyse = bool(inspection and inspection.ok and configuration_ok)

if st.button("🔍  Analyse Migration", type="primary", disabled=not ready_to_analyse):
    progress = st.progress(0.0, text="Starting the analysis…")

    def on_progress(message: str, fraction: float) -> None:
        progress.progress(min(max(fraction, 0.0), 1.0), text=message)

    analysis = analyse_migration(
        st.session_state.source_path,
        mapping_path=st.session_state.mapping_path,
        template_path=st.session_state.template_path,
        progress=on_progress,
    )
    progress.empty()
    st.session_state.analysis = analysis
    st.session_state.result = None

analysis = st.session_state.analysis

if analysis is not None and not analysis.ok:
    st.error("**The migration cannot run with these files.** Nothing was changed.")
    for message in analysis.errors:
        st.markdown(f"- {message}")
    if analysis.unresolved_mappings:
        st.markdown("**Mapping rules that could not be applied**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Mapping row": problem["mapping_row"],
                        "SAP field": problem["target_key"],
                        "Parameter": problem["target_label"],
                        "Expected legacy column": problem["legacy_label"],
                        "Legacy header row": problem["legacy_row_hint"],
                        "Reason": problem["reason"],
                    }
                    for problem in analysis.unresolved_mappings
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.info(
        "The tool stops instead of guessing which legacy column to read. Send this file and the "
        f"log `{Path(analysis.log_path).name}` to the engineering team."
    )

elif analysis is not None and analysis.ok:
    st.success(f"Analysis completed in {analysis.duration_seconds:.0f} seconds. No files were changed.")

    metrics = st.columns(6)
    metrics[0].metric("Source rows", f"{analysis.rows_processed:,}")
    metrics[1].metric("Mapping rules", f"{analysis.summary.get('active_rules', 0):,}")
    metrics[2].metric("Values migrated automatically", f"{analysis.counts.get(STATUS_OK, 0):,}")
    metrics[3].metric("Needs review", f"{analysis.review_count:,}")
    metrics[4].metric("Warnings", f"{analysis.warning_count:,}")
    metrics[5].metric("Errors", f"{analysis.error_count:,}")

    st.markdown(
        f"<span class='aks-note'>{analysis.values_to_migrate:,} values will be written into "
        f"target rows {analysis.summary.get('first_target_row')}–{analysis.summary.get('last_target_row')}. "
        f"{analysis.counts.get(STATUS_SKIPPED, 0):,} legacy cells are empty and are skipped.</span>",
        unsafe_allow_html=True,
    )

    for warning in analysis.warnings:
        st.warning(warning)

    if analysis.review_count or analysis.warning_count:
        st.markdown("#### Items requiring attention")
        st.caption(
            "These values have no defined migration. The tool does not invent a replacement — "
            "the engineering team decides."
        )

        breakdown = analysis.review_breakdown()
        if breakdown:
            st.dataframe(
                pd.DataFrame(breakdown)[
                    ["Status", "Parameter", "SAP Field", "Reason", "Items", "Example old values"]
                ],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Reason": st.column_config.TextColumn(width="large"),
                    "Items": st.column_config.NumberColumn(width="small"),
                },
            )

        with st.expander(f"Every affected value ({analysis.review_count + analysis.warning_count:,} rows)"):
            filters = st.multiselect(
                "Show",
                options=[STATUS_REVIEW, STATUS_WARNING],
                default=[STATUS_REVIEW, STATUS_WARNING],
                key="review_filter",
            )
            rows = analysis.table_rows(filters or [STATUS_REVIEW, STATUS_WARNING])
            st.dataframe(status_frame(rows), use_container_width=True, hide_index=True, height=430)
    else:
        st.success("Every legacy value is covered by the migration mapping. Nothing needs review.")

    with st.expander("All mapped values, including the ones that migrate cleanly"):
        preview_limit = 3000
        rows = analysis.table_rows(limit=preview_limit)
        st.caption(
            f"Showing the first {len(rows):,} of {analysis.summary.get('entries', 0):,} mapped values. "
            "The complete list is in the migration audit created in Step 4."
        )
        st.dataframe(status_frame(rows), use_container_width=True, hide_index=True, height=430)


# ---------------------------------------------------------------------------
# Step 4 — run the migration
# ---------------------------------------------------------------------------

st.markdown('<div class="aks-step">Step 4</div>', unsafe_allow_html=True)
st.subheader("Run migration")

ready_to_run = bool(analysis is not None and analysis.ok and env_ok)

if analysis is None:
    st.caption("Run the analysis first.")
elif not analysis.ok:
    st.caption("Fix the problems reported in Step 3 before migrating.")
elif not env_ok:
    st.error(
        "Microsoft Excel for Windows is required to write the `.xlsb` workbook. "
        "See the system check in the sidebar."
    )

with st.container(border=True):
    st.markdown(
        f"**Output file**  \n`{planned_output_name(st.session_state.source_name or 'legacy.xlsb')}`"
        if st.session_state.source_name
        else "**Output file**"
    )
    option_left, option_right = st.columns(2)
    highlight = option_left.checkbox(
        "Colour the migrated cells by status",
        value=True,
        help="Green: copied unchanged. Yellow: converted by a mapping rule. Red: needs review.",
    )
    replace_demo = option_right.checkbox(
        "Replace the demonstration rows in the template copy",
        value=True,
        help=(
            "The supplied AKS V2 template ships with example rows. They must be cleared from the "
            "copy before the migrated data is written. The master template is not changed."
        ),
    )
    st.markdown(
        "<span class='aks-note'>The legacy file and the AKS V2 master template are never modified. "
        "Each run writes to its own folder, so an earlier result is never overwritten.</span>",
        unsafe_allow_html=True,
    )
    if st.session_state.source_name:
        earlier = previous_outputs(st.session_state.source_name)
        if earlier:
            st.info(
                f"This legacy file has been migrated {len(earlier)} time(s) before. "
                f"The most recent result is kept in `{earlier[0].parent.name}` and stays untouched."
            )

if st.button("▶  Run Migration", type="primary", disabled=not ready_to_run) and analysis is not None:
    progress = st.progress(0.0, text="Preparing…")

    def on_progress(message: str, fraction: float) -> None:
        progress.progress(min(max(fraction, 0.0), 1.0), text=message)

    with st.spinner("Microsoft Excel is writing the migrated workbook. This can take a few minutes…"):
        result = run_migration(
            analysis,
            highlight=highlight,
            replace_template_data=replace_demo,
            progress=on_progress,
        )
    progress.empty()
    st.session_state.result = result

result = st.session_state.result


# ---------------------------------------------------------------------------
# Step 5 — results
# ---------------------------------------------------------------------------

if result is not None:
    st.divider()
    st.markdown('<div class="aks-step">Result</div>', unsafe_allow_html=True)

    if not result.ok:
        st.error("**Migration failed. Nothing was changed.**")
        for message in result.errors:
            st.markdown(f"- {message}")
        if result.log_path:
            st.info(f"Full technical details are in the log file `{Path(result.log_path).name}`.")
            with st.expander("Show the run log"):
                st.code(result.log_text or "", language="text")
            st.download_button(
                "Download the log file",
                data=Path(result.log_path).read_bytes(),
                file_name=Path(result.log_path).name,
                mime="text/plain",
            )
    else:
        st.success(f"**Migration completed** in {result.duration_seconds:.0f} seconds.")

        metrics = st.columns(5)
        metrics[0].metric("Rows processed", f"{result.rows_processed:,}")
        metrics[1].metric("Values migrated", f"{result.values_migrated:,}")
        metrics[2].metric("Warnings", f"{result.warning_count:,}")
        metrics[3].metric("Items requiring review", f"{result.review_count:,}")
        metrics[4].metric("Errors", f"{result.error_count:,}")

        verification = result.verification
        if verification is not None and verification.ok:
            st.markdown(
                pill("Output verified", "ok")
                + f"<span class='aks-note'>The generated workbook was reopened from disk: "
                f"{verification.populated_rows:,} data rows through row {verification.actual_last_row}, "
                "and the Migration_Report worksheet is present.</span>",
                unsafe_allow_html=True,
            )
        elif verification is not None:
            st.warning(
                "The workbook was created, but the automatic check found something to look at:\n"
                + "\n".join(f"- {problem}" for problem in verification.problems)
            )

        for warning in result.warnings:
            st.warning(warning)

        st.markdown("#### Download")
        downloads = st.columns(3)
        if result.workbook_path and result.workbook_path.exists():
            downloads[0].download_button(
                "📗  Migrated AKS workbook",
                data=result.workbook_path.read_bytes(),
                file_name=result.workbook_path.name,
                mime="application/vnd.ms-excel.sheet.binary.macroEnabled.12",
                use_container_width=True,
            )
            downloads[0].caption(result.workbook_path.name)
        if result.audit_path and result.audit_path.exists():
            downloads[1].download_button(
                "📋  Migration audit",
                data=result.audit_path.read_bytes(),
                file_name=result.audit_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            downloads[1].caption("Every value, before and after, with its status and reason.")
        if result.report_path and result.report_path.exists():
            downloads[2].download_button(
                "📑  Migration report",
                data=result.report_path.read_bytes(),
                file_name=result.report_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            downloads[2].caption("Timestamp, versions, totals and unresolved mappings.")

        if result.output_dir:
            st.markdown(
                f"<span class='aks-note'>All files are also saved in "
                f"<code>outputs\\{result.output_dir.name}</code>.</span>",
                unsafe_allow_html=True,
            )
            if st.button("Open the folder with the results"):
                open_in_explorer(result.output_dir)

        st.markdown("#### What to do next")
        st.markdown(
            f"""
1. Open the migrated workbook and check the highlighted cells.
2. Work through the **{result.review_count:,} items requiring review** in the migration audit —
   these are legacy values with no defined mapping. Decide the correct AKS V2 value with the
   engineering team.
3. Confirm that formulas, dropdowns, hidden sheets and macros still behave as expected.
4. Save the reviewed workbook under your controlled revision name.
"""
        )

        with st.expander("Show the run log"):
            st.code(result.log_text or "", language="text")
