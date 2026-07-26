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
from datetime import datetime
import statement_converter as sc

st.set_page_config(page_title="Bank Statement Converter", page_icon="📄", layout="centered")

st.title("📄 Bank Statement Converter")
st.write(
    "Upload a bank statement PDF and get back a cleaned Excel file — "
    "same column headers and order as your bank uses, spilled columns "
    "repaired, and a derived reference-number column pulled out of each "
    "transaction description."
)

with st.expander("🔒 What happens to my file?"):
    st.markdown(
        "- Your PDF is processed **in memory / temporary storage only**.\n"
        "- It is **never saved permanently** and is deleted immediately "
        "after your Excel file is generated — whether the conversion "
        "succeeds or fails.\n"
        "- Nothing is sent anywhere beyond this app doing the conversion."
    )

uploaded_file = st.file_uploader("Choose a PDF bank statement", type=["pdf"])

if uploaded_file is not None:
    st.write(f"**File:** {uploaded_file.name} ({uploaded_file.size / 1024:.0f} KB)")

    if st.button("Convert to Excel", type="primary"):
        progress_area = st.empty()
        log_lines = []

        def update_progress(msg):
            log_lines.append(msg)
            # Show the latest few lines so the user can see it's actively working
            progress_area.code("\n".join(log_lines[-8:]))

        try:
            with st.spinner("Converting... this can take a few seconds for large statements, longer for scanned PDFs."):
                pdf_bytes = uploaded_file.getvalue()
                xlsx_bytes, stats = sc.convert_bytes(
                    pdf_bytes, filename=uploaded_file.name, progress=update_progress
                )

            st.success(f"Done in {stats['elapsed_seconds']}s — {stats['n_transactions']} transactions converted.")

            col1, col2, col3 = st.columns(3)
            col1.metric("Transactions", stats["n_transactions"])
            col2.metric("Flagged for review", stats["n_flagged"])
            col3.metric("Months covered", stats["n_months"])

            if stats["n_dropped"]:
                st.warning(
                    f"{stats['n_dropped']} row(s) couldn't be matched to a valid date and were "
                    "moved to a NEEDS_REVIEW sheet instead of being silently dropped."
                )

            st.write(f"**Columns detected:** {', '.join(stats['columns'])} + Reference Number")

            output_name = uploaded_file.name.rsplit(".", 1)[0] + "_converted.xlsx"
            st.download_button(
                label="⬇️ Download Excel file",
                data=xlsx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        except ValueError as e:
            st.error(f"Couldn't convert this file: {e}")
            st.info(
                "This usually means the statement's column headers weren't recognized. "
                "If you can share a sample (redacted is fine), the header list can be extended."
            )
        except Exception as e:
            st.error(f"Something went wrong during conversion: {e}")

st.divider()
st.caption(f"Bank Statement Converter · {datetime.now().year}")
