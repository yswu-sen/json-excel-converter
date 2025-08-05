import json
import pandas as pd
import streamlit as st
import io  # 用於在記憶體中處理二進位資料流
from datetime import datetime

# --- Excel 格式化函式庫 ---
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ==============================================================================
# --- 功能一：基礎轉換 (已修改為 Streamlit 適用版本) ---
# ==============================================================================
def generate_basic_excel(json_data):
    """
    接收 JSON 資料，生成基本的 Excel 檔案，並以 bytes 形式返回。
    """
    df = pd.DataFrame(json_data)
    
    # 將 Excel 檔案寫入到記憶體的 buffer 中
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='申請人資料', index=False)
        worksheet = writer.sheets['申請人資料']
        # 調整欄寬
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max(max_length + 2, 10), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # .getvalue() 可以取得 buffer 中的 raw bytes
    return output_buffer.getvalue()

# ==============================================================================
# --- 功能二：專業格式化轉換 (已修改為 Streamlit 適用版本) ---
# ==============================================================================
def create_summary_sheet(wb, data, font_name, fallback_font, font_size):
    """創建摘要工作表 (您的原始邏輯，完全保留)"""
    ws = wb.create_sheet("資料摘要")
    summary_data = [
        ["項目", "數值"],
        ["總申請人數", len(data)],
        ["男性", len([d for d in data if d.get("性別") == "男"])],
        ["女性", len([d for d in data if d.get("性別") == "女"])],
        ["同意", len([d for d in data if d.get("勞動部檢核結果") == "同意"])],
        ["承辦中", len([d for d in data if d.get("勞動部檢核結果") == "承辦中"])],
        ["軟體技術開發", len([d for d in data if d.get("子領域") == "軟體技術開發"])],
        ["數位科技內容產製或擴散", len([d for d in data if d.get("子領域") == "數位科技內容產製或擴散"])],
    ]
    header_font = Font(bold=True, color="FFFFFF", name=font_name, size=font_size)
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    cell_font = Font(name=font_name, size=font_size)
    for row_idx, row in enumerate(summary_data, 1):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
            else:
                cell.font = cell_font
    # 字體備援處理
    try:
        for row in ws.iter_rows():
            for cell in row:
                if cell.font.name != font_name:
                    cell.font = Font(name=fallback_font, size=font_size, bold=cell.font.bold, color=cell.font.color)
    except: # 在某些環境下(如Streamlit雲端)，字體檢查可能會出錯，直接跳過
        pass
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15

def generate_advanced_excel(json_data):
    """
    接收 JSON 資料，生成專業格式的 Excel 檔案，並以 bytes 形式返回。
    """
    df = pd.DataFrame(json_data)
    wb = Workbook()
    ws = wb.active
    ws.title = "申請人資料"
    
    # --- 您的所有格式化邏輯，幾乎原封不動地保留 ---
    font_name = "Microsoft JhengHei" # 在雲端主機上可能沒有，但可以保留
    fallback_font = "Arial" # 使用一個通用的備援字體
    font_size = 14
    header_font = Font(bold=True, color="FFFFFF", name=font_name, size=font_size)
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_font = Font(name=font_name, size=font_size)
    cell_align = Alignment(vertical="top", wrap_text=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.font = header_font; cell.fill = header_fill; cell.alignment = header_align
            else:
                cell.font = cell_font; cell.alignment = cell_align
            cell.border = border
    
    try:
        for row in ws.iter_rows():
            for cell in row:
                if cell.font.name != font_name:
                    cell.font = Font(name=fallback_font, size=font_size, bold=cell.font.bold, color=cell.font.color)
    except:
        pass
    
    for column in ws.columns:
        max_length = 0
        for cell in column:
            if cell.value:
                length = sum(2 if ord(char) > 127 else 1 for char in str(cell.value))
                if length > max_length: max_length = length
        ws.column_dimensions[column[0].column_letter].width = min(max(max_length + 2, 8), 60)

    for row in ws.iter_rows(min_row=2):
        lines = max(str(cell.value).count('\n') + 1 for cell in row if cell.value) if any(cell.value for cell in row) else 1
        ws.row_dimensions[row[0].row].height = min(max(lines * 15, 15), 200)

    ws.freeze_panes = "A2"
    create_summary_sheet(wb, json_data, font_name, fallback_font, font_size)
    
    # 將完成的工作簿(workbook)存到記憶體的 buffer
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    
    return output_buffer.getvalue()

# ==============================================================================
# --- Streamlit App 主流程 ---
# ==============================================================================
def main():
    st.set_page_config(page_title="JSON 轉 Excel 工具", page_icon="🔄")
    st.title("🔄 JSON to Excel 轉換工具")

    # 使用 st.radio 替換 Tkinter 的按鈕
    conversion_type = st.radio(
        "請選擇要執行的轉換類型：",
        ('基礎轉換', '專業格式化轉換'),
        horizontal=True
    )

    # 使用 st.file_uploader 替換 Tkinter 的檔案選擇器
    uploaded_file = st.file_uploader(
        "請上傳您的 JSON 或 TXT 檔案",
        type=['json', 'txt']
    )

    if uploaded_file is not None:
        try:
            # 從上傳的檔案物件中讀取內容
            json_data = json.load(uploaded_file)
            st.success("檔案讀取成功！請點擊下方按鈕下載轉換結果。")

            excel_bytes = None
            file_suffix = ""

            if conversion_type == '基礎轉換':
                excel_bytes = generate_basic_excel(json_data)
                file_suffix = "basic"

            elif conversion_type == '專業格式化轉換':
                excel_bytes = generate_advanced_excel(json_data)
                file_suffix = "advanced"

            if excel_bytes:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                original_filename = os.path.splitext(uploaded_file.name)[0]
                
                st.download_button(
                    label=f"📥 點此下載【{conversion_type}】結果",
                    data=excel_bytes,
                    file_name=f"{original_filename}_{file_suffix}_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"處理檔案時發生錯誤：{e}")
            st.warning("請確認檔案內容是否為有效的 JSON 格式。")

if __name__ == "__main__":
    # 移除所有 Tkinter 啟動程式碼，直接呼叫 main()
    main()
