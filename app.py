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
        # 先讀取前 10 列來掃描
        df_temp = pd.read_excel(file, header=None, nrows=10, dtype=str)
    except Exception:
        return None
    
    header_idx = -1
    
    # 逐列檢查是否包含關鍵欄位
    for idx, row in df_temp.iterrows():
        row_values = [str(val).strip() for val in row.values]
        # 只要同一列裡面同時有這兩個關鍵字，就認定是標題列
        if '姓名' in row_values and '通訊地址' in row_values:
            header_idx = idx
            break
            
    file.seek(0)
    
    if header_idx != -1:
        return pd.read_excel(file, header=header_idx, dtype=str)
    else:
        return pd.read_excel(file, dtype=str)

def process_address(raw_address):
    """
    處理地址邏輯：
    從地址字串中提取郵遞區號 (例如: (950)臺東縣... -> 950, 臺東縣...)
    """
    if not isinstance(raw_address, str):
        return "   ", ""

    raw_address = raw_address.strip()
    
    # Regex 抓取開頭的 3碼數字，支援 (950) 或 950 格式
    match = re.match(r'^[\(（]?(\d{3})[\)）]?(.*)', raw_address)
    
    if match:
        zip_code = match.group(1)
        clean_addr = match.group(2).strip()
        return zip_code, clean_addr
    
    # 若地址沒寫郵遞區號，嘗試用關鍵字補全 (備用)
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
    """設定中西文字型"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(size)
    run.font.bold = bold

def generate_word_doc(df):
    """生成 Word 文件的核心邏輯"""
    doc = Document()
    
    # 設定版面: A4 大小，邊界全為 0 (3M 21320 規格)
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)

    # 建立表格 (2欄 x N列)
    total_items = len(df)
    rows_needed = (total_items + 1) // 2 
    table = doc.add_table(rows=rows_needed, cols=2)
    
    for index, row_data in df.iterrows():
        r = index // 2
        c = index % 2
        
        # 取得資料
        name = str(row_data.get('姓名', '')).strip()
        raw_address = str(row_data.get('通訊地址', '')).strip()
        
        if name == 'nan': name = ''
        if raw_address == 'nan': raw_address = ''
        
        # 處理資料
        zip_code, clean_address = process_address(raw_address)

        cell = table.rows[r].cells[c]
        cell.width = Cm(10.5)
        
        # 固定列高 2.97cm
        table.rows[r].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        table.rows[r].height = Cm(2.97) 
        cell.vertical_alignment = 1 # 垂直置中
        
        # 清除預設段落
        cell._element.clear_content()
        
        # --- 開始排版 (依照圖片樣式) ---
        
        # 1. 姓名行: [姓名] 君收
        p1 = cell.add_paragraph()
        p1.paragraph_format.left_indent = Cm(0.5) # 整體左邊界
        p1.paragraph_format.space_after = Pt(0)   # 段落後不留白
        if name:
            run1 = p1.add_run(f"{name} 君收")
            set_font(run1, size=14, bold=True) # 姓名加大加粗
            
        # 2. 郵遞區號行: 950 ( 950 )
        p2 = cell.add_paragraph()
        p2.paragraph_format.left_indent = Cm(0.5)
        p2.paragraph_format.space_after = Pt(0)
        # 格式：Zip ( Zip )
        run2 = p2.add_run(f"{zip_code} ( {zip_code} )")
        set_font(run2, size=12, bold=False)
        
        # 3. 地址行: 縮排顯示
        p3 = cell.add_paragraph()
        # 設定懸掛縮排/左縮排，讓地址往右縮進 (對齊圖片樣式)
        # 0.5 (基本邊界) + 0.8 (額外縮排) = 1.3 cm
        p3.paragraph_format.left_indent = Cm(1.3) 
        p3.paragraph_format.space_before = Pt(2) # 與上方稍微留點空隙
        
        run3 = p3.add_run(clean_address)
        set_font(run3, size=12, bold=False)

    # 存檔
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI 介面 ---

st.title("🏷️ 生日賀卡標籤生成器")
st.markdown("""
本工具專為 **3M 21320 (A4 2欄 x 10列)** 格式設計。
請上傳 Excel 通訊錄，程式將自動排版為標籤樣式。
""")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 使用自動標題偵測功能
        df = load_excel_with_auto_header(uploaded_file)
        
        if df is None:
            st.error("❌ 無法讀取 Excel 檔案，請確認格式。")
            st.stop()
        
        # 清理欄位名稱
        df.columns = [str(c).strip() for c in df.columns]
        
        # 檢查必要欄位
        required_cols = {'姓名', '通訊地址'}
        if not required_cols.issubset(df.columns):
            st.error(f"❌ 錯誤：找不到必要欄位！\n程式偵測到的欄位有：{list(df.columns)}\n請確認 Excel 中包含：{required_cols}")
            st.stop()
            
        # 顯示預覽
        st.success("✅ 成功讀取檔案！")
        st.subheader("📋 資料預覽")
        st.dataframe(df[['姓名', '通訊地址']].head())
        
        # 2. 生成按鈕
        if st.button("🚀 開始生成標籤", type="primary"):
            with st.spinner('正在生成 Word 檔...'):
                docx_buffer = generate_word_doc(df)
                
                # 3. 下載按鈕
                st.download_button(
                    label="📥 下載 Word 標籤檔 (.docx)",
                    data=docx_buffer,
                    file_name="生日賀卡標籤.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.info("💡 **列印提示**：請使用 Word 開啟，列印時選擇 **「實際大小」** 或 **縮放比例 100%**。")

    except Exception as e:
        st.error(f"程式發生錯誤：{e}")
