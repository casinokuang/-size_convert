import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel 尺碼全方位轉換", layout="wide")
st.title("Excel 尺碼橫排轉換 (全尺碼支援版)")

uploaded_file = st.file_uploader("上傳原始 Excel (original.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        
        pivot_col = '屬性/尺碼'
        value_col = '數量/值'
        
        # 1. 定義完整的尺碼排序清單 (依照圖片與描述整理)
        # 這裡的順序決定了 Excel 由左到右的顯示順序
        full_size_list = [
            # 數字尺碼
            '00', '0', '2', '4', '6', '8', '10', '12', '14', '16', 
            '18', '20', '22', '24', '26', '28', '30', '32', '34', '36',
            # 標準字母尺碼
            'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL',
            # 組合尺碼 (圖片中的新需求)
            'XS/S', 'M/L', 'XL/2XL', '3XL/4XL', '5XL/6XL', 
            'XS/M', 'L/2XL', '3XL/6XL',
            # 童裝與青少年
            '2T', '3T', '4T', '5T', 
            'J6', 'J8', 'J10', 'J12', 'J14', 'J16', 'J18', 'J20', 'J22', 'J24', 'J26', 'J28', 'Baby', 'Kids', 
        ]

        # 2. 清洗資料
        df[pivot_col] = df[pivot_col].astype(str).str.strip()
        df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)
        
        # 取得除了尺碼和數值以外的所有欄位作為索引
        fixed_cols = [c for c in df.columns if c not in [pivot_col, value_col]]
        for col in fixed_cols:
            df[col] = df[col].astype(str).str.replace(r'^(nan|None|NaT)$', '', regex=True).str.strip()

        # 3. Pivot 透視
        df_wide = df.pivot_table(
            index=fixed_cols, 
            columns=pivot_col, 
            values=value_col, 
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # 4. 補齊與排序
        df_wide.columns = df_wide.columns.astype(str)
        
        # 確保 full_size_list 中的每個尺碼都存在（若沒數據則補 0）
        for s in full_size_list:
            if s not in df_wide.columns:
                df_wide[s] = 0

        # 5. 最終欄位排序
        # 保留不在 full_size_list 裡的固定欄位，後面接著我們定義好的尺碼順序
        other_cols = [c for c in df_wide.columns if c not in full_size_list]
        final_order = other_cols + full_size_list
        
        # 只取出最終需要的欄位 (排除掉可能在 pivot 出現但不在 list 裡的雜質)
        df_final = df_wide[final_order]

        # 6. 介面與下載
        st.success(f"轉換完成！已處理 {len(df_final)} 行資料，包含組合尺碼與童裝系列。")
        st.dataframe(df_final)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 下載完整版 Excel",
            data=output.getvalue(),
            file_name="converted_all_sizes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"執行出錯：{e}")
