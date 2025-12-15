import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.table import WD_ROW_HEIGHT_RULE
from io import BytesIO
import re

# --- 設定頁面資訊 ---
st.set_page_config(
    page_title="生日賀卡標籤生成器",
    page_icon="🏷️",
    layout="centered"
)

# --- 輔助函式 ---

def load_excel_with_auto_header(file):
    """
    自動偵測 Excel 的標題列位置。
    解決第一列是標題名稱(如: 臺東縣...)而不是欄位名稱的問題。
    """
    try:
        # 讀取前 20 列來搜尋標題
        df_temp = pd.read_excel(file, header=None, nrows=20, dtype=str)
    except Exception:
        return None
    
    header_idx = -1
    
    # 逐列檢查是否包含關鍵欄位
    for idx, row in df_temp.iterrows():
        # 將整列轉為字串搜尋，移除空白
        row_values = [str(val).strip() for val in row.values]
        
        # 只要同一列裡面同時有這兩個關鍵字，就認定是標題列
        if '姓名' in row_values and '通訊地址' in row_values:
            header_idx = idx
            break
            
    # 重設檔案指標
    file.seek(0)
    
    # 根據找到的索引重新讀取
    if header_idx != -1:
        # header=header_idx 告訴 pandas 這行才是標題
        return pd.read_excel(file, header=header_idx, dtype=str)
    else:
        # 找不到就用預設
        return pd.read_excel(file, dtype=str)

def process_address(raw_address):
    """
    處理地址邏輯：提取郵遞區號與地址
    """
    if not isinstance(raw_address, str):
        return "   ", ""

    raw_address = raw_address.strip()
    
    # 抓取開頭的 3碼數字，例如 (950) 或 950
    match = re.match(r'^[\(（]?(\d{3})[\)）]?(.*)', raw_address)
    
    if match:
        zip_code = match.group(1)
        clean_addr = match.group(2).strip()
        return zip_code, clean_addr
    
    # 備用：若地址沒寫數字，用地名對照
    zip_map = {
        "花蓮市": "970", "新城鄉": "971", "秀林鄉": "972",
        "吉安鄉": "973", "壽豐鄉": "974", "鳳林鎮": "975",
        "光復鄉": "976", "豐濱鄉": "977", "瑞穗鄉": "978",
        "萬榮鄉": "979", "玉里鎮": "981", "卓溪鄉": "982",
        "富里鄉": "983", "臺東市": "950"
    }
    
    for town, code in zip_map.items():
        if town in raw_address:
            return code, raw_address
            
    return "   ", raw_address

def set_font(run, size=12, bold=False):
    """設定中西文字型 (標楷體 + Times New Roman)"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = bold

def generate_word_doc(df):
    """生成 Word 文件的核心邏輯"""
    doc = Document()
    
    # 設定版面: A4 大小，邊界全為 0 (無邊界列印)
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    # 建立表格 (2欄 x N列)
    total_items = len(df)
    
    # 計算需要多少列
    rows_needed = (total_items + 1) // 2 
    
    table = doc.add_table(rows=rows_needed, cols=2)
    
    for index, row_data in df.iterrows():
        r = index // 2
        c = index % 2
        
        # 確保 Excel 欄位名稱去除空白後讀取
        name = str(row_data.get('姓名', '')).strip()
        raw_address = str(row_data.get('通訊地址', '')).strip()
        
        if name == 'nan': name = ''
        if raw_address == 'nan': raw_address = ''
        
        # 處理地址
        zip_code, clean_address = process_address(raw_address)

        cell = table.rows[r].cells[c]
        cell.width = Cm(10.5) # 寬度固定為 A4 一半
        
        # --- 關鍵修正：高度設定為 8 列模式 ---
        # A4 高度 29.7cm / 8 = 3.7125 cm
        # 這會讓 8 個標籤剛好填滿一頁
        table.rows[r].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        table.rows[r].height = Cm(29.7 / 8) 
        
        cell.vertical_alignment = 1 # 垂直置中
        cell._element.clear_content() # 清除預設內容
        
        # --- 排版內容 (依照您的圖片樣式) ---
        
        # 1. 姓名行
        p1 = cell.add_paragraph()
        p1.paragraph_format.left_indent = Cm(0.5) # 左邊界
        p1.paragraph_format.space_after = Pt(0)
        if name:
            run1 = p1.add_run(f"{name} 君收")
            set_font(run1, size=14, bold=True)
            
        # 2. 郵遞區號行：格式 950 ( 950 )
        p2 = cell.add_paragraph()
        p2.paragraph_format.left_indent = Cm(0.5)
        p2.paragraph_format.space_after = Pt(0)
        run2 = p2.add_run(f"{zip_code} ( {zip_code} )")
        set_font(run2, size=12, bold=False)
        
        # 3. 地址行 (縮排對齊)
        p3 = cell.add_paragraph()
        p3.paragraph_format.left_indent = Cm(1.3) # 懸掛縮排效果
        p3.paragraph_format.space_before = Pt(2)
        
        run3 = p3.add_run(clean_address)
        set_font(run3, size=12, bold=False)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---

st.title("🏷️ 生日賀卡標籤生成器")
st.markdown("""
本工具專為 **A4 (2欄 x 8列)** 格式設計（每頁 16 張標籤，每張高度約 3.7cm）。
請上傳 Excel 通訊錄，程式將自動排版。
""")

# 1. 檔案上傳
uploaded_file = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 使用新的讀取函式 (自動跳過標題行)
        df = load_excel_with_auto_header(uploaded_file)
        
        if df is None:
            st.error("❌ 無法讀取 Excel 檔案，請確認格式。")
            st.stop()
        
        # 清理欄位名稱 (移除前後空白)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 檢查欄位
        required_cols = {'姓名', '通訊地址'}
        if not required_cols.issubset(df.columns):
            st.error(f"❌ 欄位錯誤！\n偵測到的欄位：{list(df.columns)}\n請確認 Excel 中包含：{required_cols}")
            # 顯示前幾行讓使用者除錯
            st.write("程式讀取到的前 5 行資料：")
            st.dataframe(df.head())
            st.stop()
            
        st.success(f"✅ 檔案讀取成功！共 {len(df)} 筆資料")
        
        # 2. 生成按鈕
        if st.button("🚀 開始生成標籤 (2x8 格式)", type="primary"):
            with st.spinner('正在生成 Word 檔...'):
                docx_buffer = generate_word_doc(df)
                
                # 3. 下載按鈕
                st.download_button(
                    label="📥 下載 Word 標籤檔 (.docx)",
                    data=docx_buffer,
                    file_name="生日賀卡標籤_2x8.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.info("💡 **列印提示**：標籤已設為填滿整頁。列印時請務必選擇 **「實際大小」 (Actual Size)**，切勿縮放。")

    except Exception as e:
        st.error(f"程式發生錯誤：{e}")

