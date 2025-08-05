import streamlit as st
import pandas as pd
import json

# --- 頁面基本設定 ---
st.set_page_config(
    page_title="JSON 轉 Excel 小工具",
    page_icon="🔄"
)

# --- 主標題 ---
st.title("🔄 JSON 轉 Excel 自動化工具")
st.write("上傳您的 JSON 檔案，即可轉換成可下載的 Excel 檔案。")

# --- 檔案上傳元件 ---
uploaded_file = st.file_uploader("請選擇一個 JSON 檔案", type=["json"])

if uploaded_file is not None:
    # 當使用者上傳檔案後執行的程式碼
    st.success(f"檔案上傳成功：{uploaded_file.name}")

    try:
        # 讀取 JSON 檔案內容
        # 注意：uploaded_file 是一個類檔案物件，需要用 .read() 來讀取
        json_data = json.load(uploaded_file)

        # --- 核心轉換邏輯 ---
        # 假設 JSON 是一個物件列表，這也是最常見的格式
        # 您可以根據您實際的 JSON 結構修改這部分
        df = pd.DataFrame(json_data)

        st.write("### 轉換結果預覽：")
        st.dataframe(df)

        # --- 將 DataFrame 轉換為 Excel ---
        # 我們將 Excel 檔案存在記憶體中，而不是實體檔案
        @st.cache_data
        def convert_df_to_excel(dataframe):
            output = pd.ExcelWriter('output.xlsx', engine='openpyxl')
            dataframe.to_excel(output, index=False, sheet_name='Sheet1')
            output.close()
            with open('output.xlsx', 'rb') as f:
                return f.read()

        excel_data = convert_df_to_excel(df)


        # --- 提供下載按鈕 ---
        st.download_button(
            label="📥 點此下載 Excel 檔案",
            data=excel_data,
            file_name=f"converted_{uploaded_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"轉換時發生錯誤：{e}")
        st.warning("請確認您的 JSON 檔案格式是否正確。")