import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_LINE_SPACING
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
    """
    try:
        # 讀取前 20 列來搜尋標題
        df_temp = pd.read_excel(file, header=None, nrows=20, dtype=str)
    except Exception:
        return None
    
    header_idx = -1
    
    # 逐列檢查是否包含關鍵欄位
    for idx, row in df_temp.iterrows():
        row_values = [str(val).strip() for val in row.values]
        if '姓名' in row_values and '通訊地址' in row_values:
            header_idx = idx
            break
            
    file.seek(0)
    
    if header_idx != -1:
        return pd.read_excel(file, header=header_idx, dtype=str)
    else:
        return pd.read_excel(file, dtype=str)

def process_address(raw_address):
    """處理地址邏輯：提取郵遞區號與地址"""
    if not isinstance(raw_address, str):
        return "   ", ""

    raw_address = raw_address.strip()
    # 抓取開頭的 3碼數字，例如 (950) 或 950
    match = re.match(r'^[\(（]?(\d{3})[\)）]?(.*)', raw_address)
    
    if match:
        zip_code = match.group(1)
        clean_addr = match.group(2).strip()
        return zip_code, clean_addr
    
    # 備用對照表
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
    
    # --- 1. 版面設定：A4 滿版零邊界 ---
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0)

    # 建立表格 (2欄 x N列)
    total_items = len(df)
    rows_needed = (total_items + 1) // 2 
    
    table = doc.add_table(rows=rows_needed, cols=2)
    table.style = 'Table Grid' # 保留格線方便查看
    
    # --- 2. 關鍵修正：強制寬度填滿 ---
    # 關閉自動調整，強制設定為固定寬度
    table.autofit = False 
    table.allow_autofit = False
    
    # 強制設定每一欄的寬度為 10.5cm (A4的一半)
    # 這會確保表格向右延伸直到紙張邊緣
    for col in table.columns:
        col.width = Cm(10.5)

    # 計算每列高度 (3.7cm * 8 = 29.6cm)
    row_height_val = Cm(3.7)

    # --- 3. 填入資料 (保留防錯機制) ---
    for i, (index, row_data) in enumerate(df.iterrows()):
        r = i // 2
        c = i % 2
        
        name = str(row_data.get('姓名', '')).strip()
        raw_address = str(row_data.get('通訊地址', '')).strip()
        
        if name == 'nan': name = ''
        if raw_address == 'nan': raw_address = ''
        
        zip_code, clean_address = process_address(raw_address)

        # 取得儲存格
        cell = table.rows[r].cells[c]
        
        # 再次確保儲存格寬度 (雙重保險)
        cell.width = Cm(10.5)
        
        # 設定高度
        table.rows[r].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        table.rows[r].height = row_height_val
        
        cell.vertical_alignment = 1 # 垂直置中
        cell._element.clear_content()
        
        # --- 排版內容 ---
        
        # 1. 姓名行
        p1 = cell.add_paragraph()
        p1.paragraph_format.left_indent = Cm(0.5)
        p1.paragraph_format.space_before = Pt(5)
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        if name:
            run1 = p1.add_run(f"{name} 君收")
            set_font(run1, size=14, bold=True)
            
        # 2. 郵遞區號行
        p2 = cell.add_paragraph()
        p2.paragraph_format.left_indent = Cm(0.5)
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        run2 = p2.add_run(f"{zip_code} ( {zip_code} )")
        set_font(run2, size=12, bold=False)
        
        # 3. 地址行
        p3 = cell.add_paragraph()
        p3.paragraph_format.left_indent = Cm(1.3)
        p3.paragraph_format.space_before = Pt(2)
        p3.paragraph_format.space_after = Pt(0)
        p3.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        run3 = p3.add_run(clean_address)
        set_font(run3, size=12, bold=False)

    # --- 4. 縮小最後游標 ---
    try:
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.paragraph_format.space_after = Pt(0)
        last_paragraph.paragraph_format.line_spacing = Pt(0)
        run = last_paragraph.add_run()
        run.font.size = Pt(1)
    except:
        pass

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---

st.title("🏷️ 生日賀卡標籤生成器")
st.markdown("""
本工具設定為 **A4 滿版 (2欄 x 8列)**。
**保證填滿整張紙張寬度 (21cm)，不再留白。**
""")

uploaded_file = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = load_excel_with_auto_header(uploaded_file)
        
        if df is None:
            st.error("❌ 無法讀取 Excel 檔案，請確認格式。")
            st.stop()
        
        df.columns = [str(c).strip() for c in df.columns]
        
        required_cols = {'姓名', '通訊地址'}
        if not required_cols.issubset(df.columns):
            st.error(f"❌ 缺少必要欄位！需包含：{required_cols}")
            st.stop()
            
        st.success(f"✅ 讀取成功！共 {len(df)} 筆資料")
        
        if st.button("🚀 生成標籤 (全寬滿版)", type="primary"):
            with st.spinner('正在生成...'):
                docx_buffer = generate_word_doc(df)
                
                st.download_button(
                    label="📥 下載 Word 標籤檔 (.docx)",
                    data=docx_buffer,
                    file_name="標籤_2x8_全滿版.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.info("💡 **列印提示**：請選擇 **「實際大小 (Actual Size)」**。")

    except Exception as e:
        st.error(f"程式發生錯誤：{e}")
        st.exception(e)
