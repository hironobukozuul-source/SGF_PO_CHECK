import streamlit as st
import pandas as pd
import datetime
import io

# Reuse your core logic from the original script
def Gen_PM_BOM(plan_data, CU_data_, DU_data_):
    abc = pd.DataFrame()
    for i in range(len(plan_data)):
        current_row = plan_data.iloc[[i], :].copy()
        current_row['Component Number'] = current_row['Material Code']
        tmp = current_row['Material Code'].values[0]
        abc = pd.concat([abc, current_row])
        
        # DU Logic
        tmp_ = DU_data_[(DU_data_["Parent material number"] == tmp) & 
                        (DU_data_['Component Description'].str.contains("OUTER", na=False))].copy()
        
        if not tmp_.empty:
            tmp_["Necessary Quantity"] = (current_row["Volume(pcs)"].values[0] / tmp_["Parent Material Quantity"]).round()
            tmp_["Material Code"] = tmp
            abc = pd.concat([abc, tmp_])
    
    return abc

# Streamlit UI Setup
st.set_page_config(page_title="SAP PO Comparison Tool", layout="wide")
st.title("📦 資材PO確認用 Web Edition")
st.write("Upload your SAP files below to generate the comparison Excel.")

# File Uploaders
col1, col2 = st.columns(2)
with col1:
    cu_file = st.file_uploader("Upload CU List (Excel)", type=["xlsx"])
    p_plan_file = st.file_uploader("Upload Old Plan (Text/CSV)", type=["txt", "csv"])

with col2:
    du_file = st.file_uploader("Upload DU List (Excel)", type=["xlsx"])
    n_plan_file = st.file_uploader("Upload New Plan (Text/CSV)", type=["txt", "csv"])

if st.button("Generate Comparison"):
    if all([cu_file, du_file, p_plan_file, n_plan_file]):
        try:
            # Load data directly from memory
            CU_data = pd.read_excel(cu_file)
            DU_data = pd.read_excel(du_file)
            
            def process_plan(file):
                # Reading the uploaded text file
                data = pd.read_csv(file, sep="\t", header=None)
                data.columns = ["Material Code","Plant Code","Production Start","Volume(pcs)","Line","Production End","Unit"]
                mapping = DU_data.set_index("Parent material number")["Parent Material Description"].to_dict()
                data["Product Code"] = data["Material Code"].map(mapping).fillna("Unknown")
                return data

            prev_bom = Gen_PM_BOM(process_plan(p_plan_file), CU_data, DU_data)
            new_bom = Gen_PM_BOM(process_plan(n_plan_file), CU_data, DU_data)
            
            # Join and compare
            idx = ["Material Code","Product Code","Production Start","Component Number"]
            prev_bom.set_index(idx, inplace=True)
            new_bom.set_index(idx, inplace=True)
            comparison = prev_bom.join(new_bom, lsuffix='_OLD', rsuffix='_NEW', how='outer').fillna(0)

            # Convert to Excel in memory for download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                comparison.to_excel(writer, sheet_name='Comparison')
            
            st.success("✅ Comparison Ready!")
            st.download_button(
                label="Download Comparison Excel",
                data=output.getvalue(),
                file_name=f"{datetime.datetime.now().strftime('%Y%m%d')}_PO_Comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Please upload all 4 required files.")
