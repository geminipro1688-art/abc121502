import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt, RGBColor
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
    table.style = 'Table Grid' # 加入格線，確保看得到邊界
    table.autofit = False 
    table.allow_autofit = False

    # --- 2. 關鍵高度計算 ---
    # A4 高度 29.7。為了避免第8行被踢走，我們設為 3.7 cm
    # 3.7 * 8 = 29.6 cm，剩下 0.1 cm 作為緩衝，這能解決「變成7張」的問題
    row_height_val = Cm(3.7) 

    for index, row_data in df.iterrows():
        r = index // 2
        c = index % 2
        
        name = str(row_data.get('姓名', '')).strip()
        raw_address = str(row_data.get('通訊地址', '')).strip()
        
        if name == 'nan': name = ''
        if raw_address == 'nan': raw_address = ''
        
        zip_code, clean_address = process_address(raw_address)

        # 取得儲存格
        cell = table.rows[r].cells[c]
        
        # --- 3. 嚴格設定寬度與高度 ---
        cell.width = Cm(10.5) # A4 寬度一半，填滿側邊
        table.rows[r].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        table.rows[r].height = row_height_val
        
        cell.vertical_alignment = 1 # 垂直置中
        cell._element.clear_content()
        
        # --- 排版內容 ---
        
        # 1. 姓名行
        p1 = cell.add_paragraph()
        p1.paragraph_format.left_indent = Cm(0.5)
        p1.paragraph_format.space_before = Pt(5) # 稍微往下壓一點
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

    # --- 4. 終極防護：縮小最後一個段落 ---
    # 這是解決「多出一頁空白頁」或「表格跑版」的關鍵
    # 把文件最後一個 Enter 鍵縮小到 1pt，讓它不會佔位子
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.paragraph_format.space_after = Pt(0)
    last_paragraph.paragraph_format.line_spacing = Pt(0)
    run = last_paragraph.add_run()
    run.font.size = Pt(1) 

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---

st.title("🏷️ 生日賀卡標籤生成器")
st.markdown("""
本工具設定為 **A4 滿版 (2欄 x 8列)**。
修正了「只有7張」的問題，現在應能剛好填滿一頁 16 張。
""")

uploaded_file = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = load_excel_with_auto_header(uploaded_file)
        
        if df is None:
            st.error("❌ 無法讀取 Excel 檔案。")
            st.stop()
        
        df.columns = [str(c).strip() for c in df.columns]
        
        required_cols = {'姓名', '通訊地址'}
        if not required_cols.issubset(df.columns):
            st.error(f"❌ 缺少必要欄位！需包含：{required_cols}")
            st.dataframe(df.head())
            st.stop()
            
        st.success(f"✅ 讀取成功！共 {len(df)} 筆資料")
        
        if st.button("🚀 生成標籤 (完美8列版)", type="primary"):
            with st.spinner('正在生成...'):
                docx_buffer = generate_word_doc(df)
                
                st.download_button(
                    label="📥 下載 Word 標籤檔 (.docx)",
                    data=docx_buffer,
                    file_name="標籤_2x8_滿版修正.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.warning("⚠️ **列印重要提示**：")
                st.markdown("""
                1. 開啟 Word 檔後，若看到最後一行有一點點空白是正常的（為了防止跑版）。
                2. 列印時請選擇 **「實際大小 (Actual Size)」**。
                3. 請確認印表機設定中的邊界已歸零，或使用「無邊界列印」功能。
                """)

    except Exception as e:
        st.error(f"錯誤：{e}")
