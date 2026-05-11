"""
O'Neill Institute Health Care Litigation Tracker — Case Processing Tool
Streamlit web app with Georgetown Law visual style.
"""

import fitz
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Litigation Tracker Tools",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Paths ───────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent
DOCKET_DIR    = BASE_DIR / "TrialCourtDockets"
COMPLAINT_DIR = BASE_DIR / "TrialCourtComplaints"
OUTPUT_DIR    = BASE_DIR / "outputs"
META_FILE     = BASE_DIR / "output_meta.json"

JUDGE_CSV          = BASE_DIR / "Federal Judicial Center Export.csv"
GOALS_CSV          = BASE_DIR / "GoalsMapping.csv"
GOALS_EXAMPLES     = BASE_DIR / "GoalsExamples.xlsx"
ISSUES_CSV         = BASE_DIR / "IssuesMapping.csv"
ISSUES_LIST_CSV    = BASE_DIR / "Issues.csv"
ISSUES_EXAMPLES    = BASE_DIR / "LegalIssuesExamples.xlsx"
ANALYSIS_EXAMPLES  = BASE_DIR / "AnalysisExamples.xlsx"
DISTRICT_PDF       = BASE_DIR / "28 USC Ch5 District Courts.pdf"

for d in [DOCKET_DIR, COMPLAINT_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

def _get_api_key() -> str:
    try:
        return st.secrets.get("OPENROUTER_API_KEY", "") or ""
    except Exception:
        return os.environ.get("OPENROUTER_API_KEY", "")

# ── Metadata helpers ────────────────────────────────────────────────────────────
def load_meta() -> dict:
    if META_FILE.exists():
        try:
            return json.loads(META_FILE.read_text())
        except Exception:
            return {}
    return {}

def save_meta(meta: dict):
    META_FILE.write_text(json.dumps(meta, indent=2))

# ── CSS ─────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Merriweather:ital,wght@0,400;0,700;1,400&family=Source+Sans+3:wght@300;400;600;700&display=swap');

/* ── Global reset ── */
html, body, [class*="css"] {
    font-family: 'Source Sans 3', sans-serif;
    color: #1a1a1a;
}

/* ── Page background ── */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main {
    background-color: #e3e9f0 !important;
}

/* ── Hide sidebar and go full-width ── */
[data-testid="stSidebar"]                { display: none !important; }
[data-testid="stSidebarCollapsedControl"] { display: none !important; }

/* Zero out Streamlit's top padding across all known version selectors */
[data-testid="stMainBlockContainer"],
[data-testid="stMain"] > div,
.main .block-container,
.block-container {
    padding-top: 0 !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
}

/* ── Georgetown-style top header ── */
.gt2-header {
    background: #e3e9f0;
    display: flex;
    align-items: stretch;
    justify-content: space-between;
    border-bottom: 2px solid #fff;
    margin: 3.5rem -2rem 0 -2rem;
    padding: 0 2.5rem;
}
.gt2-branding {
    display: flex;
    align-items: center;
    gap: 0.8rem;
    padding: 1rem 0;
}
.gt2-logo-mark {
    font-size: 1.5rem;
    line-height: 1;
}
.gt2-title {
    font-family: 'Merriweather', Georgia, serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: #002147;
    line-height: 1.25;
}
.gt2-subtitle {
    font-size: 0.7rem;
    color: #555;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-top: 2px;
}
.gt2-nav {
    display: flex;
    align-items: stretch;
}
.gt2-nav-link {
    color: #002147 !important;
    text-decoration: none !important;
    font-weight: 600;
    font-size: 0.875rem;
    padding: 0 1rem;
    border-bottom: 3px solid transparent;
    display: flex;
    align-items: center;
    transition: background 0.15s, border-color 0.15s;
    white-space: nowrap;
}
.gt2-nav-link:hover {
    background: rgba(0, 33, 71, 0.07);
    border-bottom-color: #16c5f9;
    color: #002147 !important;
    text-decoration: none !important;
}
.gt2-nav-active {
    border-bottom-color: #002147 !important;
}

/* ── Page heading ── */
.gt2-page-header {
    padding: 1.8rem 0 1.4rem 0;
    border-bottom: 1px solid #e8e6e0;
    margin-bottom: 1.5rem;
}
.gt2-eyebrow {
    font-size: 0.78rem;
    text-transform: uppercase;
    font-weight: 700;
    letter-spacing: 0.1em;
    color: #888;
    margin-bottom: 0.35rem;
}
.gt2-page-title {
    font-family: 'Merriweather', Georgia, serif;
    font-size: 2rem;
    font-weight: 700;
    color: #002147;
    margin: 0;
    line-height: 1.2;
}

/* ── Sub-section headings ── */
.gt-section-title {
    font-family: 'Merriweather', Georgia, serif;
    font-size: 1.25rem;
    font-weight: 700;
    color: #002147;
    border-bottom: 2px solid #c8a951;
    padding-bottom: 0.45rem;
    margin-bottom: 1.1rem;
}
.gt-section-sub {
    font-size: 0.88rem;
    color: #555;
    margin-top: -0.7rem;
    margin-bottom: 1.2rem;
}

/* ── Cards ── */
.gt-card {
    background: #f8f7f4;
    border: 1px solid #ddd;
    border-left: 4px solid #002147;
    border-radius: 4px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.8rem;
}
.gt-card-gold {
    border-left-color: #c8a951;
}

/* ── Buttons ── */
.stButton > button {
    background: #002147 !important;
    color: white !important;
    border: none !important;
    border-radius: 3px !important;
    font-family: 'Source Sans 3', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.04em !important;
    text-transform: uppercase !important;
    font-size: 0.82rem !important;
    padding: 0.55rem 1.3rem !important;
    transition: background 0.2s !important;
}
.stButton > button:hover {
    background: #c8a951 !important;
    color: #002147 !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"],
[data-testid="stFileUploaderDropzone"] {
    border: 1px solid #002147 !important;
    border-radius: 4px !important;
    background: #fffdf6 !important;
}
[data-testid="stFileUploaderDropzone"]:hover,
[data-testid="stFileUploaderDropzone"]:focus-within {
    border-color: #16c5f9 !important;
}
[data-testid="stFileUploaderProgressBar"] > div,
[data-testid="stFileUploaderProgressBar"] > div > div {
    background-color: #16c5f9 !important;
}

/* ── Tabs ── */
[data-baseweb="tab-highlight"] {
    background-color: #16c5f9 !important;
}
[data-baseweb="tab"][aria-selected="true"] {
    color: #002147 !important;
}
[data-baseweb="tab"]:hover {
    color: #002147 !important;
}

/* ── Tables ── */
.gt-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.9rem;
}
.gt-table th {
    background: #002147;
    color: white;
    padding: 0.6rem 0.9rem;
    text-align: left;
    font-family: 'Source Sans 3', sans-serif;
    font-weight: 600;
    letter-spacing: 0.03em;
    font-size: 0.82rem;
    text-transform: uppercase;
}
.gt-table td {
    padding: 0.55rem 0.9rem;
    border-bottom: 1px solid #e8e6e0;
}
.gt-table tr:nth-child(even) td {
    background: #f8f7f4;
}
.gt-table tr:hover td {
    background: #fff8e6;
}

/* ── Status badges ── */
.badge-ok   { background:#1a6b3c; color:white; padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:600; }
.badge-warn { background:#c8a951; color:#002147; padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:600; }
.badge-err  { background:#9b1c1c; color:white; padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:600; }

/* ── Input fields ── */
.stTextInput > div > div > input {
    border: 1px solid #ccc !important;
    border-radius: 3px !important;
    font-family: 'Source Sans 3', sans-serif !important;
}
.stTextInput > div > div > input:focus {
    border-color: #002147 !important;
    box-shadow: 0 0 0 2px rgba(0,33,71,0.15) !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    font-family: 'Source Sans 3', sans-serif !important;
    font-weight: 600 !important;
    color: #002147 !important;
    background: #f8f7f4 !important;
}

/* ── File table ── */
.ft-head {
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: #002147;
    padding-bottom: 0.3rem;
}

/* ── Divider ── */
hr { border-color: #e0ddd5; }

/* ── Info boxes ── */
[data-testid="stInfo"] {
    background: #eef2f8 !important;
    border-left: 4px solid #002147 !important;
    border-radius: 3px !important;
}
[data-testid="stSuccess"] {
    border-left: 4px solid #1a6b3c !important;
    border-radius: 3px !important;
}
[data-testid="stWarning"] {
    border-left: 4px solid #c8a951 !important;
    border-radius: 3px !important;
}
[data-testid="stError"] {
    border-left: 4px solid #9b1c1c !important;
    border-radius: 3px !important;
}
</style>
""", unsafe_allow_html=True)

# ── Page routing via query params ───────────────────────────────────────────────
_VALID_PAGES = {"process", "output", "judicial", "districts", "inputs"}
page = st.query_params.get("page", "process")
if page not in _VALID_PAGES:
    page = "process"

def _nav_link(label: str, key: str) -> str:
    active = "gt2-nav-active" if page == key else ""
    return f'<a href="?page={key}" target="_self" class="gt2-nav-link {active}">{label}</a>'

# ── Georgetown-style top header ──────────────────────────────────────────────────
st.markdown(f"""
<div class="gt2-header">
  <div class="gt2-branding">
    <span class="gt2-logo-mark">⚖️</span>
    <div>
      <div class="gt2-title">Litigation Information Extraction Tool</div>
      <div class="gt2-subtitle">Health Care Litigation Tracker</div>
    </div>
  </div>
  <nav class="gt2-nav">
    {_nav_link("Process Cases", "process")}
    {_nav_link("Output Files", "output")}
    {_nav_link("Judicial Data", "judicial")}
    {_nav_link("District Courts", "districts")}
    {_nav_link("Model Inputs", "inputs")}
  </nav>
</div>
""", unsafe_allow_html=True)

# ── Shared helper: render a CSV or XLSX file management section ─────────────────
def _render_file_section(
    path: Path,
    display_name: str,
    upload_key: str,
    file_types: list[str],
    description: str = "",
    is_xlsx: bool = False,
):
    if path.exists():
        mtime = datetime.fromtimestamp(path.stat().st_mtime)
        size_kb = path.stat().st_size / 1024
        st.markdown(f"""
        <div class="gt-card">
          <strong>{display_name}</strong><br>
          <span style="color:#555; font-size:0.88rem;">
            Last modified: {mtime.strftime('%B %d, %Y at %I:%M %p')} &nbsp;·&nbsp; {size_kb:,.1f} KB
          </span>
        </div>
        """, unsafe_allow_html=True)

        try:
            if is_xlsx:
                df = pd.read_excel(path, engine="openpyxl")
            else:
                df = pd.read_csv(path, encoding="utf-8-sig")

            col_dl, col_spacer = st.columns([1, 3])
            with col_dl:
                with open(path, "rb") as fh:
                    ext = "xlsx" if is_xlsx else "csv"
                    mime = (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if is_xlsx else "text/csv"
                    )
                    st.download_button(
                        f"⬇ Download {ext.upper()}",
                        data=fh,
                        file_name=display_name,
                        mime=mime,
                        use_container_width=True,
                        key=f"dl_{upload_key}",
                    )

            st.markdown(f"**{len(df):,} rows · {len(df.columns)} columns**")

            if description:
                st.caption(description)

            search = st.text_input("🔍 Search", "", key=f"search_{upload_key}")
            if search:
                mask = df.apply(
                    lambda col: col.astype(str).str.contains(search, case=False, na=False)
                ).any(axis=1)
                df_display = df[mask]
                st.caption(f"{len(df_display):,} matching rows")
            else:
                df_display = df

            st.dataframe(df_display, use_container_width=True, height=420)

        except Exception as exc:
            st.error(f"Could not read {display_name}: {exc}")
    else:
        st.warning(f"{display_name} not found.")

    st.markdown("---")
    st.markdown("**Replace / Update File**")
    ext_label = "XLSX" if is_xlsx else "CSV"
    st.caption(f"Upload a new version of the {ext_label} to replace the existing file.")

    uploaded = st.file_uploader(
        f"Upload updated {display_name}",
        type=file_types,
        key=upload_key,
        label_visibility="collapsed",
    )
    if uploaded:
        path.write_bytes(uploaded.read())
        size_kb = uploaded.size / 1024
        st.success(f"✓ {display_name} updated ({size_kb:.1f} KB)")
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1 — Process Cases
# ══════════════════════════════════════════════════════════════════════════════
if page == "process":
    st.markdown("""
    <div class="gt2-page-header">
      <div class="gt2-eyebrow">Case Processing Tool</div>
      <h1 class="gt2-page-title">Process Cases</h1>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    # ── Docket uploads ──────────────────────────────────────────────────────────
    with col1:
        st.markdown("**Step 1 — Upload Docket PDFs**")
        docket_files = st.file_uploader(
            "Docket PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key="docket_uploader",
            label_visibility="collapsed",
        )
        if docket_files:
            for f in docket_files:
                dest = DOCKET_DIR / f.name
                dest.write_bytes(f.read())
            st.success(f"✓  {len(docket_files)} docket file(s) staged.")

        staged_dockets = list(DOCKET_DIR.glob("*.pdf"))
        if staged_dockets:
            with st.expander(f"Staged dockets ({len(staged_dockets)})", expanded=False):
                for p in staged_dockets:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"📄 {p.name}")
                    if c2.button("✕", key=f"del_d_{p.name}", help="Remove"):
                        p.unlink()
                        st.rerun()

    # ── Complaint uploads ───────────────────────────────────────────────────────
    with col2:
        st.markdown("**Step 2 — Upload Complaint PDFs**")
        complaint_files = st.file_uploader(
            "Complaint PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key="complaint_uploader",
            label_visibility="collapsed",
        )
        if complaint_files:
            for f in complaint_files:
                dest = COMPLAINT_DIR / f.name
                dest.write_bytes(f.read())
            st.success(f"✓  {len(complaint_files)} complaint file(s) staged.")

        staged_complaints = list(COMPLAINT_DIR.glob("*.pdf"))
        if staged_complaints:
            with st.expander(f"Staged complaints ({len(staged_complaints)})", expanded=False):
                for p in staged_complaints:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"📄 {p.name}")
                    if c2.button("✕", key=f"del_c_{p.name}", help="Remove"):
                        p.unlink()
                        st.rerun()

    st.markdown("---")

    # ── API key status ──────────────────────────────────────────────────────────
    api_key = _get_api_key()
    if api_key:
        st.markdown(
            "**OpenRouter API Key** &nbsp;"
            '<span class="badge-ok">Configured</span>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "**OpenRouter API Key** &nbsp;"
            '<span class="badge-err">Not configured</span>',
            unsafe_allow_html=True,
        )
        st.caption(
            "No API key found. Goals, Issues, Potential Impact, and Why This Matters will "
            "return 'Not identified'. To enable AI classification, add "
            "`OPENROUTER_API_KEY = \"sk-or-...\"` to `.streamlit/secrets.toml` "
            "(local) or the Secrets panel in Streamlit Cloud."
        )

    st.markdown("---")

    # ── Run processing ──────────────────────────────────────────────────────────
    st.markdown("**Step 3 — Name output and process**")

    col_name, col_btn = st.columns([3, 1], gap="medium")
    with col_name:
        output_name = st.text_input(
            "Output file name (without extension)",
            value=f"Tracker Data Summary {date.today().strftime('%m.%d.%y')}",
            placeholder="e.g. Q2 2025 Cases",
        )
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        run = st.button("▶  Process", use_container_width=True)

    if run:
        if not list(DOCKET_DIR.glob("*.pdf")):
            st.error("No docket PDFs staged. Please upload at least one docket before processing.")
        elif not output_name.strip():
            st.error("Please provide an output file name.")
        else:
            with st.spinner("Running summarization pipeline — this may take several minutes…"):
                try:
                    run_env = {**os.environ}
                    if api_key:
                        run_env["OPENROUTER_API_KEY"] = api_key

                    result = subprocess.run(
                        [sys.executable, str(BASE_DIR / "summarize_cases.py")],
                        cwd=str(BASE_DIR),
                        capture_output=True,
                        text=True,
                        timeout=600,
                        env=run_env,
                    )

                    today_str = date.today().strftime("%m.%d.%y")
                    generated = list(BASE_DIR.glob(f"Tracker Data Summary {today_str}*.*"))

                    if not generated and result.returncode != 0:
                        st.error("Processing failed.")
                        if result.stderr:
                            st.code(result.stderr, language="text")
                    else:
                        meta = load_meta()
                        now_ts = datetime.now().isoformat(timespec="seconds")

                        for gen_path in generated:
                            suffix = gen_path.suffix
                            new_name = f"{output_name.strip()}{suffix}"
                            dest = OUTPUT_DIR / new_name
                            counter = 1
                            while dest.exists():
                                dest = OUTPUT_DIR / f"{output_name.strip()} ({counter}){suffix}"
                                counter += 1
                            shutil.move(str(gen_path), str(dest))
                            meta[dest.name] = {
                                "created": now_ts,
                                "source_dockets": [p.name for p in DOCKET_DIR.glob("*.pdf")],
                            }

                        save_meta(meta)
                        st.success(f"✓ Processing complete! {len(generated)} file(s) saved to Output Files.")

                        if result.stdout:
                            with st.expander("Pipeline output log", expanded=False):
                                st.code(result.stdout, language="text")
                        if result.stderr:
                            with st.expander("Warnings / stderr", expanded=False):
                                st.code(result.stderr, language="text")

                except subprocess.TimeoutExpired:
                    st.error("Processing timed out after 10 minutes.")
                except Exception as exc:
                    st.error(f"Unexpected error: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2 — Output Files
# ══════════════════════════════════════════════════════════════════════════════
elif page == "output":
    st.markdown("""
    <div class="gt2-page-header">
      <div class="gt2-eyebrow">Generated Data</div>
      <h1 class="gt2-page-title">Output Files</h1>
    </div>
    """, unsafe_allow_html=True)

    output_files = sorted(OUTPUT_DIR.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)
    meta = load_meta()

    if not output_files:
        st.info("No output files yet. Use the **Process Cases** page to generate files.")
    else:
        # ── Table header ────────────────────────────────────────────────────────
        h = st.columns([5, 3, 1, 1, 1, 1])
        h[0].markdown('<div class="ft-head">File</div>', unsafe_allow_html=True)
        h[1].markdown('<div class="ft-head">Created</div>', unsafe_allow_html=True)
        h[2].markdown('<div class="ft-head">Size</div>', unsafe_allow_html=True)
        st.markdown(
            '<hr style="margin:0 0 0.25rem 0; border-color:#002147; border-width:2px;">',
            unsafe_allow_html=True,
        )

        # ── One row per file ────────────────────────────────────────────────────
        for f in output_files:
            created_str = meta.get(f.name, {}).get("created", "—")
            try:
                created_display = datetime.fromisoformat(created_str).strftime("%b %d, %Y · %I:%M %p")
            except Exception:
                created_display = created_str

            size_kb = f.stat().st_size / 1024

            row = st.columns([5, 3, 1, 1, 1, 1])
            with row[0]:
                st.markdown(f"**{f.name}**")
            with row[1]:
                st.markdown(
                    f'<span style="color:#555; font-size:0.88rem;">{created_display}</span>',
                    unsafe_allow_html=True,
                )
            with row[2]:
                st.markdown(
                    f'<span style="color:#555; font-size:0.88rem;">{size_kb:.1f} KB</span>',
                    unsafe_allow_html=True,
                )
            with row[3]:
                with open(f, "rb") as fh:
                    st.download_button(
                        "⬇", data=fh, file_name=f.name,
                        key=f"dl_{f.name}", use_container_width=True,
                    )
            with row[4]:
                if st.button("✏", key=f"ren_btn_{f.name}", use_container_width=True):
                    st.session_state[f"renaming_{f.name}"] = True
            with row[5]:
                if st.button("🗑", key=f"del_btn_{f.name}", use_container_width=True):
                    st.session_state[f"confirm_del_{f.name}"] = True

            st.markdown(
                '<hr style="margin:0.25rem 0; border-color:#e8e6e0;">',
                unsafe_allow_html=True,
            )

            if st.session_state.get(f"renaming_{f.name}"):
                new_stem = st.text_input(
                    "New name (without extension)",
                    value=f.stem,
                    key=f"new_name_{f.name}",
                )
                c1, c2 = st.columns(2)
                if c1.button("Save", key=f"save_ren_{f.name}"):
                    new_path = OUTPUT_DIR / f"{new_stem.strip()}{f.suffix}"
                    if new_path.exists():
                        st.error("A file with that name already exists.")
                    else:
                        f.rename(new_path)
                        if f.name in meta:
                            meta[new_path.name] = meta.pop(f.name)
                            save_meta(meta)
                        st.session_state.pop(f"renaming_{f.name}", None)
                        st.rerun()
                if c2.button("Cancel", key=f"cancel_ren_{f.name}"):
                    st.session_state.pop(f"renaming_{f.name}", None)
                    st.rerun()

            if st.session_state.get(f"confirm_del_{f.name}"):
                st.warning(f"Delete **{f.name}**? This cannot be undone.")
                c1, c2 = st.columns(2)
                if c1.button("Yes, delete", key=f"yes_del_{f.name}"):
                    f.unlink()
                    if f.name in meta:
                        del meta[f.name]
                        save_meta(meta)
                    st.session_state.pop(f"confirm_del_{f.name}", None)
                    st.rerun()
                if c2.button("Cancel", key=f"no_del_{f.name}"):
                    st.session_state.pop(f"confirm_del_{f.name}", None)
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3 — Judicial Data
# ══════════════════════════════════════════════════════════════════════════════
elif page == "judicial":
    st.markdown("""
    <div class="gt2-page-header">
      <div class="gt2-eyebrow">Federal Judicial Center</div>
      <h1 class="gt2-page-title">Judicial Data</h1>
    </div>
    """, unsafe_allow_html=True)

    col_info, col_btn = st.columns([3, 1])
    with col_info:
        if JUDGE_CSV.exists():
            mtime = datetime.fromtimestamp(JUDGE_CSV.stat().st_mtime)
            size_kb = JUDGE_CSV.stat().st_size / 1024
            st.markdown(f"""
            <div class="gt-card">
              <strong>Federal Judicial Center Export.csv</strong><br>
              <span style="color:#555; font-size:0.88rem;">
                Last updated: {mtime.strftime('%B %d, %Y at %I:%M %p')} &nbsp;·&nbsp; {size_kb:,.1f} KB
              </span>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.warning("Judicial data file not found.")

    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        update = st.button("Update Judicial Data", use_container_width=True)

    if update:
        with st.spinner("Downloading latest FJC judge data…"):
            try:
                result = subprocess.run(
                    [sys.executable, str(BASE_DIR / "update_judge_data.py")],
                    cwd=str(BASE_DIR),
                    capture_output=True,
                    text=True,
                    timeout=180,
                )
                if result.returncode == 0:
                    st.success("✓  Judicial data updated successfully.")
                else:
                    st.error("Update failed.")
                    if result.stderr:
                        st.code(result.stderr, language="text")
            except subprocess.TimeoutExpired:
                st.error("Update timed out.")
            except Exception as exc:
                st.error(f"Error: {exc}")

    st.markdown("---")

    if JUDGE_CSV.exists():
        try:
            df = pd.read_csv(JUDGE_CSV, encoding="latin-1", low_memory=False)
            st.markdown("**Preview of Federal Judicial Center Export:**")
            st.markdown(f"**{len(df):,} records · {len(df.columns)} columns**")

            search = st.text_input("🔍 Filter by judge name or court", "")
            if search:
                mask = df.apply(
                    lambda col: col.astype(str).str.contains(search, case=False, na=False)
                ).any(axis=1)
                df = df[mask]
                st.caption(f"{len(df):,} matching records")

            st.dataframe(df, use_container_width=True, height=480)

            with open(JUDGE_CSV, "rb") as fh:
                st.download_button(
                    "⬇ Download CSV",
                    data=fh,
                    file_name="Federal Judicial Center Export.csv",
                    mime="text/csv",
                )
        except Exception as exc:
            st.error(f"Could not read CSV: {exc}")
    else:
        st.info("Click **Update Judicial Data** to download the latest FJC export.")


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4 — District Courts PDF
# ══════════════════════════════════════════════════════════════════════════════
elif page == "districts":
    st.markdown("""
    <div class="gt2-page-header">
      <div class="gt2-eyebrow">28 U.S.C. Chapter 5</div>
      <h1 class="gt2-page-title">District Courts</h1>
    </div>
    """, unsafe_allow_html=True)

    if DISTRICT_PDF.exists():
        size_kb = DISTRICT_PDF.stat().st_size / 1024
        mtime = datetime.fromtimestamp(DISTRICT_PDF.stat().st_mtime)

        st.markdown(f"""
        <div class="gt-card">
          <strong>28 USC Ch5 District Courts.pdf</strong><br>
          <span style="color:#555; font-size:0.88rem;">
            On file since: {mtime.strftime('%B %d, %Y')} &nbsp;·&nbsp; {size_kb:,.1f} KB
          </span>
        </div>
        """, unsafe_allow_html=True)

        col_dl, col_spacer = st.columns([1, 3])
        with col_dl:
            with open(DISTRICT_PDF, "rb") as fh:
                st.download_button(
                    "⬇ Download PDF",
                    data=fh,
                    file_name="28 USC Ch5 District Courts.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
    else:
        st.warning("District Courts PDF not found on the backend.")

    st.markdown("---")
    st.markdown("**Replace / Update File**")
    st.caption("Upload a new version of the PDF to replace the existing file.")

    new_pdf = st.file_uploader(
        "Upload updated District Courts PDF",
        type=["pdf"],
        key="district_pdf_upload",
        label_visibility="collapsed",
    )
    if new_pdf:
        DISTRICT_PDF.write_bytes(new_pdf.read())
        st.success(f"✓ File updated: {new_pdf.name} ({new_pdf.size / 1024:.1f} KB)")
        st.rerun()

    if DISTRICT_PDF.exists():
        st.markdown("---")
        try:
            st.markdown("**28 U.S.C. Chapter 5 District Courts Preview:**")
            doc = fitz.open(str(DISTRICT_PDF))
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                st.image(pix.tobytes("png"), use_container_width=True)
            doc.close()
        except Exception as exc:
            st.error(f"Could not render PDF: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 5 — Model Inputs
# ══════════════════════════════════════════════════════════════════════════════
elif page == "inputs":
    st.markdown("""
    <div class="gt2-page-header">
      <div class="gt2-eyebrow">Classification and References for AI-Generated Fields</div>
      <h1 class="gt2-page-title">Model Inputs</h1>
    </div>
    """, unsafe_allow_html=True)

    tab_goals_map, tab_goals_ex, tab_issues_list, tab_issues_map, tab_issues_ex, tab_analysis_ex = st.tabs([
        "Goals Mapping",
        "Goals Examples",
        "Issues List",
        "Issues Mapping",
        "Issues Examples",
        "Analysis Examples",
    ])

    with tab_goals_map:
        _render_file_section(
            path=GOALS_CSV,
            display_name="GoalsMapping.csv",
            upload_key="goals_csv_upload",
            file_types=["csv"],
            is_xlsx=False,
        )

    with tab_goals_ex:
        _render_file_section(
            path=GOALS_EXAMPLES,
            display_name="GoalsExamples.xlsx",
            upload_key="goals_xlsx_upload",
            file_types=["xlsx"],
            is_xlsx=True,
        )

    with tab_issues_list:
        _render_file_section(
            path=ISSUES_LIST_CSV,
            display_name="Issues.csv",
            upload_key="issues_list_csv_upload",
            file_types=["csv"],
            description="Canonical list of valid legal issue names used to validate and normalize pipeline output.",
            is_xlsx=False,
        )

    with tab_issues_map:
        _render_file_section(
            path=ISSUES_CSV,
            display_name="IssuesMapping.csv",
            upload_key="issues_csv_upload",
            file_types=["csv"],
            is_xlsx=False,
        )

    with tab_issues_ex:
        _render_file_section(
            path=ISSUES_EXAMPLES,
            display_name="LegalIssuesExamples.xlsx",
            upload_key="issues_xlsx_upload",
            file_types=["xlsx"],
            is_xlsx=True,
        )

    with tab_analysis_ex:
        _render_file_section(
            path=ANALYSIS_EXAMPLES,
            display_name="AnalysisExamples.xlsx",
            upload_key="analysis_xlsx_upload",
            file_types=["xlsx"],
            is_xlsx=True,
        )
