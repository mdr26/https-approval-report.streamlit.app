"""
app.py — Approval Report Generator
Run with: streamlit run app.py
"""

import os
import tempfile
from datetime import datetime

import streamlit as st
from generate_report import generate_report

st.set_page_config(
    page_title="Approval Report Generator",
    page_icon="🏥",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #f5f4f0; }
#MainMenu, footer, header { visibility: hidden; }

.top-banner {
    background: #1a2e44;
    color: white;
    padding: 28px 36px 22px;
    border-radius: 16px;
    margin-bottom: 28px;
}
.top-banner h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    font-weight: 400;
    margin: 0 0 4px;
    color: white;
}
.top-banner p { font-size: 0.9rem; color: #a8bdd0; margin: 0; }

.step-card {
    background: white;
    border-radius: 14px;
    padding: 20px 24px;
    margin-bottom: 12px;
    border: 1.5px solid #e8e4dc;
}
.step-number {
    display: inline-block;
    background: #1a2e44;
    color: white;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 1px;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 8px;
    text-transform: uppercase;
}
.step-title { font-size: 1rem; font-weight: 600; color: #1a2e44; margin-bottom: 3px; }
.step-desc  { font-size: 0.85rem; color: #7a8a9a; margin: 0; }

.metric-row { display: flex; gap: 10px; margin: 18px 0; flex-wrap: wrap; }
.metric-card {
    flex: 1; min-width: 90px;
    background: white;
    border-radius: 12px;
    padding: 14px 10px;
    text-align: center;
    border: 1.5px solid #e8e4dc;
}
.metric-card.total    { border-top: 4px solid #1a2e44; }
.metric-card.approved { border-top: 4px solid #2ecc71; }
.metric-card.partial  { border-top: 4px solid #f39c12; }
.metric-card.pending  { border-top: 4px solid #3498db; }
.metric-card.rejected { border-top: 4px solid #e74c3c; }
.metric-num   { font-family: 'DM Serif Display', serif; font-size: 1.9rem; color: #1a2e44; line-height: 1; margin-bottom: 4px; }
.metric-label { font-size: 0.75rem; color: #7a8a9a; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }

.dup-box    { background: #fffbf0; border: 1.5px solid #f0d080; border-radius: 10px; padding: 14px 18px; font-size: 0.87rem; color: #7a5c00; margin: 12px 0; }
.no-dup-box { background: #f0faf4; border: 1.5px solid #90d4aa; border-radius: 10px; padding: 14px 18px; font-size: 0.87rem; color: #1a6e3c; margin: 12px 0; }
.success-banner { background: #f0faf4; border: 1.5px solid #90d4aa; border-radius: 12px; padding: 14px 20px; color: #1a6e3c; font-size: 0.92rem; font-weight: 500; margin: 12px 0; }
.error-box  { background: #fff5f5; border: 1.5px solid #f5c0c0; border-radius: 10px; padding: 16px 20px; color: #a02020; font-size: 0.88rem; margin: 12px 0; }
.error-box strong { display: block; margin-bottom: 6px; font-size: 1rem; }

[data-testid="stFileUploader"] { background: white; border-radius: 12px; padding: 8px; border: 1.5px dashed #c8d4e0; }

.stButton > button {
    background: #1a2e44 !important; color: white !important;
    border: none !important; border-radius: 10px !important;
    padding: 14px 28px !important; font-size: 1rem !important;
    font-weight: 500 !important; width: 100% !important;
}
.stButton > button:hover { background: #243d58 !important; }

[data-testid="stDownloadButton"] > button {
    background: #2ecc71 !important; color: white !important;
    border: none !important; border-radius: 10px !important;
    padding: 14px 28px !important; font-size: 1rem !important;
    font-weight: 600 !important; width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="top-banner">
    <h1>Approval Report Generator</h1>
    <p>Insurance Approval Summary &nbsp;·&nbsp; Daily Report Tool</p>
</div>
""", unsafe_allow_html=True)

# ── Steps ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="step-card">
    <div class="step-number">Step 1</div>
    <div class="step-title">Export the file from the insurance portal</div>
    <p class="step-desc">Log in and export the latest Approval Summary as an Excel file (.xlsx).</p>
</div>
<div class="step-card">
    <div class="step-number">Step 2</div>
    <div class="step-title">Upload the file below</div>
    <p class="step-desc">Click the upload area and select the file you just exported.</p>
</div>
<div class="step-card">
    <div class="step-number">Step 3</div>
    <div class="step-title">Generate and Download</div>
    <p class="step-desc">Click Generate Report. When it finishes, click Download to save the formatted report.</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("---")
uploaded_file = st.file_uploader(
    "📂  Upload your Approval Summary file here",
    type=["xlsx"],
    help="Select the Excel file exported from the insurance portal.",
)

# ── Generate ──────────────────────────────────────────────────────────────────
if uploaded_file is not None:
    st.markdown(f"""
    <div class="success-banner">✅ &nbsp; File ready: <strong>{uploaded_file.name}</strong></div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("⚡  Generate Report", type="primary", use_container_width=True):
        with st.spinner("Building your report… this usually takes under 10 seconds."):
            try:
                ts = datetime.now().strftime("%d%m%Y_%H%M%S")
                output_filename = f"Approval_Report_{ts}.xlsx"

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    output_path = tmp.name

                stats = generate_report(uploaded_file, output_path)

                with open(output_path, "rb") as f:
                    report_bytes = f.read()
                os.unlink(output_path)

                st.markdown("---")
                st.markdown("### ✅ Report Ready")

                # Metrics
                st.markdown(f"""
                <div class="metric-row">
                    <div class="metric-card total">
                        <div class="metric-num">{stats['total']}</div>
                        <div class="metric-label">Total</div>
                    </div>
                    <div class="metric-card approved">
                        <div class="metric-num">{stats['approved']}</div>
                        <div class="metric-label">Approved</div>
                    </div>
                    <div class="metric-card partial">
                        <div class="metric-num">{stats['partial']}</div>
                        <div class="metric-label">Partial</div>
                    </div>
                    <div class="metric-card pending">
                        <div class="metric-num">{stats['pending']}</div>
                        <div class="metric-label">Pending</div>
                    </div>
                    <div class="metric-card rejected">
                        <div class="metric-num">{stats['rejected']}</div>
                        <div class="metric-label">Rejected</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Deduplication notice
                if stats["rows_removed"] > 0:
                    st.markdown(f"""
                    <div class="dup-box">
                        ℹ️ &nbsp;<strong>{stats['rows_removed']} duplicate rows were automatically removed.</strong><br>
                        One row is kept per unique Approval Number + Status combination.
                        (e.g. 3 Approved rows for same number → kept as 1)
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="no-dup-box">✅ &nbsp; No duplicate rows found.</div>
                    """, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                st.download_button(
                    label="📥  Download Report",
                    data=report_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption(f"Saved as: {output_filename}")

            except ValueError as ve:
                import pandas as pd
                st.markdown(f"""
                <div class="error-box">
                    <strong>❌ Could not read the file</strong>{ve}
                </div>
                """, unsafe_allow_html=True)
                try:
                    uploaded_file.seek(0)
                    xl = pd.ExcelFile(uploaded_file)
                    st.info(
                        f"Sheets found in your file: **{', '.join(xl.sheet_names)}**\n\n"
                        "Please make sure you are uploading the correct Approval Summary export file."
                    )
                except Exception:
                    pass

            except Exception as e:
                st.markdown(f"""
                <div class="error-box">
                    <strong>❌ Something went wrong</strong>{str(e)}<br><br>
                    Please check you are uploading the correct file and try again.
                    If the problem continues, contact your administrator.
                </div>
                """, unsafe_allow_html=True)
else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👆  Please upload the Approval Summary file above to get started.")

st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;color:#aaa;font-size:0.8rem;'>"
    "Approval Report Generator &nbsp;·&nbsp; For internal use only</div>",
    unsafe_allow_html=True
)
