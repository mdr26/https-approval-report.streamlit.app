"""
app.py — Approval Report Generator
A clean, professional Streamlit app designed for daily use.
Run with: streamlit run app.py
"""

import os
import tempfile
from datetime import datetime

import streamlit as st
from generate_report import generate_report

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Approval Report Generator",
    page_icon="🏥",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Background */
.stApp {
    background: #f5f4f0;
}

/* Hide streamlit branding */
#MainMenu, footer, header { visibility: hidden; }

/* Top banner */
.top-banner {
    background: #1a2e44;
    color: white;
    padding: 28px 36px 22px 36px;
    border-radius: 16px;
    margin-bottom: 28px;
}
.top-banner h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    font-weight: 400;
    margin: 0 0 4px 0;
    color: white;
    letter-spacing: -0.5px;
}
.top-banner p {
    font-size: 0.95rem;
    color: #a8bdd0;
    margin: 0;
    font-weight: 300;
}

/* Step cards */
.step-card {
    background: white;
    border-radius: 14px;
    padding: 24px 28px;
    margin-bottom: 16px;
    border: 1.5px solid #e8e4dc;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.step-number {
    display: inline-block;
    background: #1a2e44;
    color: white;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 1px;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 10px;
    text-transform: uppercase;
}
.step-title {
    font-size: 1.05rem;
    font-weight: 600;
    color: #1a2e44;
    margin-bottom: 4px;
}
.step-desc {
    font-size: 0.88rem;
    color: #7a8a9a;
    margin: 0;
}

/* Metric cards */
.metric-row {
    display: flex;
    gap: 12px;
    margin: 20px 0;
    flex-wrap: wrap;
}
.metric-card {
    flex: 1;
    min-width: 100px;
    background: white;
    border-radius: 12px;
    padding: 16px 14px;
    text-align: center;
    border: 1.5px solid #e8e4dc;
}
.metric-card.approved { border-top: 4px solid #2ecc71; }
.metric-card.partial  { border-top: 4px solid #f39c12; }
.metric-card.pending  { border-top: 4px solid #3498db; }
.metric-card.rejected { border-top: 4px solid #e74c3c; }
.metric-card.total    { border-top: 4px solid #1a2e44; }
.metric-num {
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    color: #1a2e44;
    line-height: 1;
    margin-bottom: 4px;
}
.metric-label {
    font-size: 0.78rem;
    color: #7a8a9a;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Dup warning */
.dup-box {
    background: #fffbf0;
    border: 1.5px solid #f0d080;
    border-radius: 10px;
    padding: 14px 18px;
    font-size: 0.88rem;
    color: #7a5c00;
    margin: 12px 0;
}
.no-dup-box {
    background: #f0faf4;
    border: 1.5px solid #90d4aa;
    border-radius: 10px;
    padding: 14px 18px;
    font-size: 0.88rem;
    color: #1a6e3c;
    margin: 12px 0;
}

/* Download section */
.download-box {
    background: #1a2e44;
    border-radius: 14px;
    padding: 24px 28px;
    text-align: center;
    margin-top: 20px;
}
.download-box p {
    color: #a8bdd0;
    font-size: 0.88rem;
    margin: 0 0 14px 0;
}

/* File uploader styling */
[data-testid="stFileUploader"] {
    background: white;
    border-radius: 12px;
    padding: 8px;
    border: 1.5px dashed #c8d4e0;
}

/* Button */
.stButton > button {
    background: #1a2e44 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 14px 28px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 500 !important;
    width: 100% !important;
    cursor: pointer !important;
    transition: background 0.2s !important;
}
.stButton > button:hover {
    background: #243d58 !important;
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: #2ecc71 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 14px 28px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    width: 100% !important;
}

/* Error box */
.error-box {
    background: #fff5f5;
    border: 1.5px solid #f5c0c0;
    border-radius: 10px;
    padding: 16px 20px;
    color: #a02020;
    font-size: 0.9rem;
    margin: 12px 0;
}
.error-box strong { display: block; margin-bottom: 6px; font-size: 1rem; }

/* Success banner */
.success-banner {
    background: #f0faf4;
    border: 1.5px solid #90d4aa;
    border-radius: 12px;
    padding: 16px 20px;
    color: #1a6e3c;
    font-size: 0.95rem;
    font-weight: 500;
    margin: 12px 0;
}
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="top-banner">
    <h1>Approval Report Generator</h1>
    <p>Insurance Approval Summary — Daily Report Tool &nbsp;·&nbsp; Al Mouwasat Hospital</p>
</div>
""", unsafe_allow_html=True)


# ── Step 1: Instructions ──────────────────────────────────────────────────────
st.markdown("""
<div class="step-card">
    <div class="step-number">Step 1</div>
    <div class="step-title">Export the Approval Summary from the insurance portal</div>
    <p class="step-desc">Log in to the insurance system and export the latest Approval Summary as an Excel file (.xlsx).</p>
</div>

<div class="step-card">
    <div class="step-number">Step 2</div>
    <div class="step-title">Upload the file below</div>
    <p class="step-desc">Click the upload area below and select the file you just exported. The app accepts any .xlsx file.</p>
</div>

<div class="step-card">
    <div class="step-number">Step 3</div>
    <div class="step-title">Click Generate, then Download</div>
    <p class="step-desc">Hit the Generate Report button. Once it finishes, click Download to save the formatted report to your computer.</p>
</div>
""", unsafe_allow_html=True)


# ── Step 2: File Upload ───────────────────────────────────────────────────────
st.markdown("---")
uploaded_file = st.file_uploader(
    "📂  Upload your Approval Summary file here",
    type=["xlsx"],
    help="Select the Excel file exported from the insurance portal.",
    label_visibility="visible",
)

# ── Step 3: Generate ──────────────────────────────────────────────────────────
if uploaded_file is not None:
    st.markdown(f"""
    <div class="success-banner">
        ✅ &nbsp; File ready: <strong>{uploaded_file.name}</strong>
    </div>
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

                # ── Results ───────────────────────────────────────────────
                st.markdown("---")
                st.markdown("### ✅ Report Ready")
                st.markdown(f"*Source sheet detected: `{stats['detected_sheet']}`*")

                # Metric cards
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

                # Duplicate notice
                if stats["duplicates"] > 0:
                    st.markdown(f"""
                    <div class="dup-box">
                        ⚠️ &nbsp;<strong>{stats['duplicates']} rows have a duplicate Approval Number.</strong>
                        These rows are highlighted in yellow in the <em>All Approval Summary</em> sheet so they are easy to spot.
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="no-dup-box">
                        ✅ &nbsp; No duplicate Approval Numbers found.
                    </div>
                    """, unsafe_allow_html=True)

                # Download
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥  Download Report",
                    data=report_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption(f"File will be saved as: {output_filename}")

            except ValueError as ve:
                import pandas as pd
                st.markdown(f"""
                <div class="error-box">
                    <strong>❌ Could not read the file</strong>
                    {ve}
                </div>
                """, unsafe_allow_html=True)
                try:
                    uploaded_file.seek(0)
                    xl = pd.ExcelFile(uploaded_file)
                    st.info(
                        f"📋 Sheets found in your file: **{', '.join(xl.sheet_names)}**\n\n"
                        "The app needs a sheet containing these columns: "
                        "`CoCode`, `CoName`, `ApprovedStatus`, `ApprovalNum`, `ResponseRemarks`.\n\n"
                        "Please make sure you are uploading the correct file."
                    )
                except Exception:
                    pass

            except Exception as e:
                st.markdown(f"""
                <div class="error-box">
                    <strong>❌ Something went wrong</strong>
                    {str(e)}<br><br>
                    Please check that you are uploading the correct Approval Summary file
                    and try again. If the problem continues, contact your system administrator.
                </div>
                """, unsafe_allow_html=True)

else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👆  Please upload the Approval Summary file above to get started.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center; color:#aaa; font-size:0.8rem;'>"
    "Approval Report Generator &nbsp;·&nbsp; For internal use only"
    "</div>",
    unsafe_allow_html=True
)
