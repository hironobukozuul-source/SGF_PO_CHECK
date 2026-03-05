import streamlit as st
import pandas as pd
import datetime
import io

# 画面設定
st.set_page_config(page_title="SAP製造指示照合ツール", page_icon="📦", layout="wide")

def Gen_PM_BOM(plan_data, CU_data_, DU_data_):
    """BOM生成ロジック"""
    abc = pd.DataFrame()
    for i in range(len(plan_data)):
        current_row = plan_data.iloc[[i], :].copy()
        current_row['Component Number'] = current_row['Material Code']
        tmp = current_row['Material Code'].values[0]
        abc = pd.concat([abc, current_row])
        
        # DUロジック (外装)
        tmp_ = DU_data_[(DU_data_["Parent material number"] == tmp) & 
                        (DU_data_['Component Description'].str.contains("OUTER", na=False))].copy()
        
        if not tmp_.empty:
            qty_val = current_row["Volume(pcs)"].values[0] / tmp_["Parent Material Quantity"].values[0]
            tmp_["Necessary Quantity"] = round(qty_val)
            tmp_["Material Code"] = tmp
            tmp_["Product Code"] = current_row['Product Code'].values[0]
            tmp_["Production Start"] = current_row['Production Start'].values[0]
            abc = pd.concat([abc, tmp_[["Component Number","Component Description","Necessary Quantity","Material Code","Product Code","Production Start"]]])
    
        # CUロジック
        cu_matches = DU_data_[(DU_data_["Parent material number"] == tmp) & 
                             (DU_data_['Component Description'].str.contains("_CU", na=False))]
        if not cu_matches.empty:
            CU_NO = cu_matches["Component Number"].iloc[0]
            tmp__ = CU_data_[(CU_data_["Parent material number"] == CU_NO) & 
                             (CU_data_["Base Unit of Measure.1"].str.contains("PC", na=False))].copy()
            
            if not tmp__.empty:
                tmp__["Necessary Quantity"] = current_row["Volume(pcs) Pach"].values[0] if "Volume(pcs) Pach" in current_row else current_row["Volume(pcs)"].values[0]
                tmp__["Material Code"] = tmp
                tmp__["Product Code"] = current_row['Product Code'].values[0]
                tmp__["Production Start"] = current_row['Production Start'].values[0]
                abc = pd.concat([abc, tmp__[["Component Number","Component Description","Necessary Quantity","Material Code","Product Code","Production Start"]]])
    return abc

st.title("📦 SAP製造指示（PO）照合システム")
st.markdown("新旧の計画データを比較し、数量に差異がある箇所を自動でハイライトします。")

# サイドバー: マスタデータ
with st.sidebar:
    st.header("1. マスタデータのアップロード")
    cu_file = st.file_uploader("CUリスト (Excel)", type=["xlsx"])
    du_file = st.file_uploader("DUリスト (Excel)", type=["xlsx"])
    st.info("SAPからエクスポートしたExcelファイルを選択してください。")

# メインエリア: 計画データ
col1, col2 = st.columns(2)
with col1:
    st.subheader("旧計画 (旧バージョン)")
    p_plan_file = st.file_uploader("旧計画ファイル (.txt)", type=["txt"])
with col2:
    st.subheader("新計画 (最新バージョン)")
    n_plan_file = st.file_uploader("新計画ファイル (.txt)", type=["txt"])

if st.button("🔍 照合レポートを作成する"):
    if all([cu_file, du_file, p_plan_file, n_plan_file]):
        try:
            with st.spinner("データを解析中..."):
                CU_data = pd.read_excel(cu_file)
                DU_data = pd.read_excel(du_file)
                
                def process_plan(file):
                    # SAPのtxtエクスポートは通常タブ区切り
                    df = pd.read_csv(file, sep="\t", header=None)
                    df.columns = ["品目コード","プラント","製造開始日","数量(pcs)","ライン","製造終了日","単位"]
                    # 英語の内部処理用カラム名に一時的に合わせる
                    df.columns = ["Material Code","Plant Code","Production Start","Volume(pcs)","Line","Production End","Unit"]
                    mapping = DU_data.set_index("Parent material number")["Parent Material Description"].to_dict()
                    df["Product Code"] = df["Material Code"].map(mapping).fillna("不明")
                    return df

                # 照合ロジック実行
                prev_bom = Gen_PM_BOM(process_plan(p_plan_file), CU_data, DU_data)
                new_bom = Gen_PM_BOM(process_plan(n_plan_file), CU_data, DU_data)
                
                idx_cols = ["Material Code","Product Code","Production Start","Component Number"]
                prev_bom.set_index(idx_cols, inplace=True)
                new_bom.set_index(idx_cols, inplace=True)
                
                # データの結合
                comparison = prev_bom.join(new_bom, lsuffix='_旧', rsuffix='_新', how='outer').fillna(0).reset_index()

                # 日本語のカラム名に変換
                jap_columns = {
                    "Material Code": "品目コード",
                    "Product Code": "製品名",
                    "Production Start": "製造開始日",
                    "Component Number": "構成品番号",
                    "Component Description_旧": "構成品名称(旧)",
                    "Necessary Quantity_旧": "必要数量(旧)",
                    "Component Description_新": "構成品名称(新)",
                    "Necessary Quantity_新": "必要数量(新)"
                }
                comparison.rename(columns=jap_columns, inplace=True)

                # 差異がある行のインデックスを取得 (Python側で判定)
                diff_indices = comparison[comparison['必要数量(旧)'] != comparison['必要数量(新)']].index.tolist()

                # Excel出力
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    comparison.to_excel(writer, sheet_name='照合結果', index=False)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['照合結果']
                    
                    # ハイライト書式 (薄い赤)
                    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    
                    # 差異行に書式を適用
                    for row_idx in diff_indices:
                        worksheet.set_row(row_idx + 1, None, red_format)

                    # 列幅の自動調整
                    for i, col in enumerate(comparison.columns):
                        worksheet.set_column(i, i, 20)

                st.success(f"✅ 照合完了: {len(diff_indices)} 件の差異が見つかりました。")
                st.download_button(
                    label="📥 照合レポート(Excel)をダウンロード",
                    data=output.getvalue(),
                    file_name=f"計画照合結果_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
    else:
        st.error("すべてのファイルをアップロードしてください。")
