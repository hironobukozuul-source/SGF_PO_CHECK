import streamlit as st
import pandas as pd
import datetime
import io

# Professional UI Styling
st.set_page_config(page_title="SAP PO Auditor", page_icon="📦", layout="wide")

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
            qty_val = current_row["Volume(pcs)"].values[0] / tmp_["Parent Material Quantity"].values[0]
            tmp_["Necessary Quantity"] = round(qty_val)
            tmp_["Material Code"] = tmp
            tmp_["Product Code"] = current_row['Product Code'].values[0]
            tmp_["Production Start"] = current_row['Production Start'].values[0]
            abc = pd.concat([abc, tmp_[["Component Number","Component Description","Necessary Quantity","Material Code","Product Code","Production Start"]]])
    
        # CU Logic
        cu_matches = DU_data_[(DU_data_["Parent material number"] == tmp) & 
                             (DU_data_['Component Description'].str.contains("_CU", na=False))]
        if not cu_matches.empty:
            CU_NO = cu_matches["Component Number"].iloc[0]
            tmp__ = CU_data_[(CU_data_["Parent material number"] == CU_NO) & 
                             (CU_data_["Base Unit of Measure.1"].str.contains("PC", na=False))].copy()
            
            if not tmp__.empty:
                tmp__["Necessary Quantity"] = current_row["Volume(pcs)"].values[0]
                tmp__["Material Code"] = tmp
                tmp__["Product Code"] = current_row['Product Code'].values[0]
                tmp__["Production Start"] = current_row['Production Start'].values[0]
                abc = pd.concat([abc, tmp__[["Component Number","Component Description","Necessary Quantity","Material Code","Product Code","Production Start"]]])
    return abc

st.title("📦 SAP PO Comparison Tool")

# Sidebar for Master Data
with st.sidebar:
    st.header("1. Master Data")
    cu_file = st.file_uploader("Upload CU List (Excel)", type=["xlsx"])
    du_file = st.file_uploader("Upload DU List (Excel)", type=["xlsx"])

# Main area
col1, col2 = st.columns(2)
with col1:
    st.subheader("Old Plan")
    p_plan_file = st.file_uploader("Select Old Plan (.txt)", type=["txt"])
with col2:
    st.subheader("New Plan")
    n_plan_file = st.file_uploader("Select New Plan (.txt)", type=["txt"])

if st.button("🔍 Generate Highlighted Comparison"):
    if all([cu_file, du_file, p_plan_file, n_plan_file]):
        try:
            with st.spinner("Processing..."):
                CU_data = pd.read_excel(cu_file)
                DU_data = pd.read_excel(du_file)
                
                def process_plan(file):
                    df = pd.read_csv(file, sep="\t", header=None)
                    df.columns = ["Material Code","Plant Code","Production Start","Volume(pcs)","Line","Production End","Unit"]
                    mapping = DU_data.set_index("Parent material number")["Parent Material Description"].to_dict()
                    df["Product Code"] = df["Material Code"].map(mapping).fillna("Unknown")
                    return df

                prev_bom = Gen_PM_BOM(process_plan(p_plan_file), CU_data, DU_data)
                new_bom = Gen_PM_BOM(process_plan(n_plan_file), CU_data, DU_data)
                
                idx_cols = ["Material Code","Product Code","Production Start","Component Number"]
                prev_bom.set_index(idx_cols, inplace=True)
                new_bom.set_index(idx_cols, inplace=True)
                
                comparison = prev_bom.join(new_bom, lsuffix='_OLD', rsuffix='_NEW', how='outer').fillna(0)

                # Export to Excel with light Red Highlighting
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    comparison.to_excel(writer, sheet_name='Comparison')
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['Comparison']
                    
                    # Format: Light Red Background
                    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    
                    max_row = len(comparison)
                    # Total columns including indices
                    max_col = len(comparison.columns) + len(comparison.index.names)

                    # Formula: Compare Column K and Column R
                    # Row 2 is the first row of data
                    worksheet.conditional_format(1, 0, max_row, max_col - 1, {
                        'type':     'formula',
                        'formula':  '=$K2<>$R2',
                        'format':   red_format
                    })

                st.success("✅ Comparison Generated!")
                st.download_button(
                    label="📥 Download Red Highlighted Report",
                    data=output.getvalue(),
                    file_name="PO_Comparison_Red_Highlights.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Logic Error: {e}")
    else:
        st.error("Missing files! Please upload all 4 required SAP exports.")
