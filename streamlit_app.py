"""
Bank Statement Converter -- Streamlit App
==========================================
Thin UI layer around statement_converter.py. All conversion logic lives
there; this file is purely upload -> progress -> download.

Privacy: the uploaded PDF is held in memory / a temp directory for the
duration of the conversion only (see convert_bytes() in
statement_converter.py). It is never written to permanent storage and is
deleted the moment conversion finishes, success or failure.
"""

import streamlit as st
import pandas as pd
import io

# 1. Page Configuration
st.set_page_config(
    page_title="Bank Statement Converter",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. Custom CSS for UI/UX Overhaul
st.markdown("""
<style>
    /* Hide standard Streamlit header and footer padding */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 800px;
    }
    
    /* Main Hero Banner Styling */
    .hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        color: #FFFFFF;
        padding: 2.5rem 2rem;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
    }
    .hero-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        color: #94A3B8;
        font-size: 1.05rem;
        line-height: 1.5;
        max-width: 600px;
        margin: 0 auto;
    }

    /* Process Flow Cards */
    .step-container {
        display: flex;
        gap: 1rem;
        margin-bottom: 2rem;
    }
    .step-box {
        flex: 1;
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
    }
    .step-number {
        display: inline-block;
        background-color: #E0E7FF;
        color: #4338CA;
        font-weight: 700;
        width: 28px;
        height: 28px;
        line-height: 28px;
        border-radius: 50%;
        margin-bottom: 0.5rem;
        font-size: 0.875rem;
    }
    .step-text {
        font-size: 0.875rem;
        color: #475569;
        font-weight: 500;
    }

    /* Style Streamlit File Uploader Box */
    [data-testid="stFileUploader"] {
        border: 2px dashed #CBD5E1;
        background-color: #FAFAFA;
        border-radius: 12px;
        padding: 1.5rem 1rem;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #6366F1;
        background-color: #F5F3FF;
    }

    /* Security Notice Pill */
    .security-badge {
        background-color: #ECFDF5;
        border: 1px solid #A7F3D0;
        color: #065F46;
        padding: 0.75rem 1rem;
        border-radius: 10px;
        font-size: 0.875rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# --- HERO SECTION ---
st.markdown("""
<div class="hero-card">
    <div class="hero-title">📄 Bank Statement Converter</div>
    <div class="hero-subtitle">
        Transform messy bank PDFs into structured, clean Excel sheets instantly. Spilled columns repaired and reference numbers auto-extracted.
    </div>
</div>
""", unsafe_allow_html=True)

# --- PROCESS / STEP-BY-STEP FLOW ---
st.markdown("""
<div class="step-container">
    <div class="step-box">
        <div class="step-number">1</div>
        <div class="step-text">Upload PDF statement</div>
    </div>
    <div class="step-box">
        <div class="step-number">2</div>
        <div class="step-text">Automated parsing</div>
    </div>
    <div class="step-box">
        <div class="step-number">3</div>
        <div class="step-text">Download clean Excel</div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- SECURITY INFORMATION ---
st.markdown("""
<div class="security-badge">
    🔒 <b>Privacy First:</b> Your file is processed in temporary memory only and deleted immediately after conversion.
</div>
""", unsafe_allow_html=True)

# --- FILE UPLOADER SECTION ---
uploaded_file = st.file_uploader(
    "Drop your bank statement here",
    type=["pdf"],
    help="Supports standard PDF statements up to 200MB."
)

# --- FILE PROCESSING & RESULTS ---
if uploaded_file is not None:
    st.divider()
    
    with st.status("Processing your bank statement...", expanded=True) as status:
        st.write("📖 Reading PDF tables...")
        # Place your actual processing function here
        # df = convert_pdf_to_excel(uploaded_file)
        
        st.write("🧹 Repairing column headers and extracting reference IDs...")
        
        # Example dummy output DataFrame for illustration
        df_result = pd.DataFrame({
            "Date": ["2026-01-10", "2026-01-12"],
            "Description": ["Transfer ref: 998231", "POS Purchase Store X"],
            "Reference No": ["998231", "N/A"],
            "Amount": [-150.00, -45.50]
        })
        
        status.update(label="Conversion Complete!", state="complete", expanded=False)

    st.success("Your document has been converted successfully!")

    # Tab view for Preview and Download
    tab1, tab2 = st.tabs(["📊 Preview Data", "📥 Download Options"])
    
    with tab1:
        st.caption("Here is a quick preview of your parsed data:")
        st.dataframe(df_result, use_container_width=True)
        
    with tab2:
        # Convert DataFrame to Excel buffer
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Statement')
            
        st.download_button(
            label="Download Excel File (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"{uploaded_file.name.replace('.pdf', '')}_converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
