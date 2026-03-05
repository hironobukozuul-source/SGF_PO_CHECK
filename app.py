import streamlit as st
import pandas as pd
import io

# Professional UI Styling
st.set_page_config(page_title="SAP PO Auditor", page_icon="📦", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("📦 SAP PO Comparison Tool")
st.info("Upload your Master Data (CU/DU) and Plan files to identify discrepancies.")

# Sidebar for Master Data
with st.sidebar:
    st.header("1. Master Data")
    cu_file = st.file_uploader("CU List (Excel)", type=["xlsx"])
    du_file = st.file_uploader("DU List (Excel)", type=["xlsx"])
    st.divider()
    st.caption("v1.2.0 - Senior Dev Edition")

# Main area for Plans
col1, col2 = st.columns(2)
with col1:
    st.subheader("Old Plan")
    p_plan_file = st.file_uploader("Select Old Plan (.txt)", type=["txt"])
with col2:
    st.subheader("New Plan")
    n_plan_file = st.file_uploader("Select New Plan (.txt)", type=["txt"])

if st.button("🔍 Generate Comparison"):
    if all([cu_file, du_file, p_plan_file, n_plan_file]):
        with st.spinner("Processing BOMs and comparing plans..."):
            # Load Data
            CU_data = pd.read_excel(cu_file)
            DU_data = pd.read_excel(du_file)
            
            # Re-using your logic (optimized)
            def process_plan(file):
                df = pd.read_csv(file, sep="\t", header=None)
                df.columns = ["Material Code","Plant","Start","Vol","Line","End","Unit"]
                mapping = DU_data.set_index("Parent material number")["Parent Material Description"].to_dict()
                df["Product Name"] = df["Material Code"].map(mapping).fillna("N/A")
                return df

            # ... [BOM Logic remains same as previous turn] ...
            # (Assuming BOM generation results in 'comparison_df')
            
            # MOCK PREVIEW FOR DEMO
            st.subheader("Data Preview")
            st.dataframe(comparison.head(10), use_container_width=True)

            # Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                comparison.to_excel(writer)
            
            st.download_button(
                label="📥 Download Comparison Report",
                data=output.getvalue(),
                file_name="PO_Comparison_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Missing files! Please upload all 4 required SAP exports.")
