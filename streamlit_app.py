import streamlit as st
import pandas as pd
import io

st.title("Excel 尺碼轉換 (尺碼版)")

uploaded_file = st.file_uploader("上傳 orginal.xlsx", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    
    pivot_col = '屬性/尺碼'
    value_col = '數量/值'
    
    # 1. 定義 A.xlsx 完整的尺碼清單 (順序 00 到 32)
    full_size_list = [
        '00', '0', '2', '4', '6', '8', '10', '12', '14', '16', 
        '18', '19', '20', '22', '24', '26', '28', '30', '32'
    ]

    # 2. 清洗固定欄位
    fixed_cols = [c for c in df.columns if c not in [pivot_col, value_col]]
    for col in fixed_cols:
        df[col] = df[col].astype(str).str.strip().replace(['nan', 'None', 'NaT'], '')
    
    df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)

    # 3. 執行 Pivot
    df_wide = df.pivot_table(
        index=fixed_cols, 
        columns=pivot_col, 
        values=value_col, 
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # 4. 【核心修復】強制檢查並補齊缺失的尺碼 (例如 00 碼)
    # 我們將所有尺碼都轉為字串來比對
    df_wide.columns = df_wide.columns.astype(str)
    for s in full_size_list:
        if s not in df_wide.columns:
            df_wide[s] = 0  # 如果數據裡沒 00 碼，就補一列 0

    # 5. 重新排列順序：固定欄位在前，00-32 尺碼在後
    # 找出除了尺碼以外的原始欄位
    actual_fixed = [c for c in df_wide.columns if c not in full_size_list]
    
    # 拼接：固定欄位 + 完整的尺碼清單
    final_order = actual_fixed + full_size_list
    
    # 過濾重複並確保欄位存在
    final_order = list(dict.fromkeys([c for c in final_order if c in df_wide.columns]))
    
    df_final = df_wide[final_order]

    # 6. 顯示與下載
    st.success("轉換完成！已補齊 00 碼並排至最後。")
    st.dataframe(df_final)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 下載 convert_restored.xlsx",
        data=output.getvalue(),
        file_name="convert_with_all_sizes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )