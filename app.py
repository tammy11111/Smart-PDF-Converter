import streamlit as st
import cv2
import numpy as np
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image
import io
import os
import re # 新增正則表達式套件

# --- 網頁設定 ---
st.set_page_config(page_title="PDF 轉 PPT (清單強化版)", layout="wide")

st.title("📄 PDF 轉 PPT：清單強化 + 智慧過濾")
st.markdown("""
**本次更新邏輯：**
1. **清單強制拆解**：只要是 Bullet Point (`•`, `1.`, `-`)，**不論顏色**一律轉為文字。
2. **干擾移除**：自動塗掉右下角的 "NotebookLM"。
3. **智慧過濾**：彩色圖表標籤保留在背景，黑色內文與標題轉為文字。
""")

# --- 參數設定 ---
OCR_LANG = 'chi_tra+eng'
TARGET_DPI = 300
BLACK_THRESHOLD = 80 

# --- 核心功能 ---

def is_list_item(text):
    """
    判斷字串是否像是一個條列式清單 (Bullet Point)
    """
    text = text.strip()
    if not text: return False
    
    # 1. 常見符號開頭
    markers = ['•', '●', '○', '▪', '▫', '◆', '◇', '➢', '➣', '➤', '→', '-', '—', '–', '*', '>']
    if any(text.startswith(m) for m in markers):
        return True
        
    # 2. 數字/字母編號開頭 (例如 "1.", "2)", "A.", "(a)")
    # 正則表達式：開頭是數字或字母，後面跟著 . 或 )
    pattern = r'^(\d+|[a-zA-Z])[\.\)]\s+'
    if re.match(pattern, text):
        return True
        
    return False

def is_text_black(image_np, x, y, w, h):
    """判斷文字是否為黑色"""
    img_h, img_w, _ = image_np.shape
    x = max(0, x); y = max(0, y)
    w = min(w, img_w - x); h = min(h, img_h - y)
    if w <= 0 or h <= 0: return False

    roi = image_np[y:y+h, x:x+w]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    _, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    if cv2.countNonZero(mask) == 0: return False

    mean_val = cv2.mean(roi, mask=mask)
    b, g, r = mean_val[0], mean_val[1], mean_val[2]
    
    if b < BLACK_THRESHOLD and g < BLACK_THRESHOLD and r < BLACK_THRESHOLD:
        return True
    return False

def get_smart_median_color(image_np, x, y, w, h):
    """區域中位數吸色"""
    img_h, img_w, _ = image_np.shape
    sample_w = 10
    
    x1 = max(0, x - sample_w)
    x2 = x
    y1 = y
    y2 = min(img_h, y + min(h, 10))
    
    if (x2 - x1) < 2:
        x1 = x
        x2 = min(img_w, x + sample_w)
        y1 = max(0, y - 5)
        y2 = y
        
    try:
        roi = image_np[y1:y2, x1:x2]
        if roi.size == 0: return (255, 255, 255)
        median_color = np.median(roi, axis=(0, 1))
        return (int(median_color[0]), int(median_color[1]), int(median_color[2]))
    except:
        return (255, 255, 255)

def get_font_size_float(heights_px):
    """計算字體大小"""
    if not heights_px: return 12.0
    avg_h_px = np.mean(heights_px)
    size_pt = (avg_h_px / TARGET_DPI) * 72 * 0.85
    if size_pt < 9: size_pt = 10
    if size_pt > 120: size_pt = 120
    return size_pt

def process_pdf(uploaded_file):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    bytes_data = uploaded_file.getvalue()
    
    status_text = st.empty()
    progress_bar = st.progress(0)
    
    status_text.text("正在轉檔與分析 (300 DPI)...")
    images = convert_from_bytes(bytes_data, dpi=TARGET_DPI)
    total_pages = len(images)
    
    for i, img in enumerate(images):
        status_text.text(f"🔄 正在處理第 {i+1} / {total_pages} 頁...")
        
        img_np = np.array(img)
        img_np = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        img_h, img_w, _ = img_np.shape
        
        # 1. 執行 OCR
        data = pytesseract.image_to_data(img, lang=OCR_LANG, output_type=Output.DICT)
        
        paragraphs = {}
        n_boxes = len(data['text'])
        
        clean_bg_img = img_np.copy()
        
        # --- 第一階段：收集資料 ---
        for j in range(n_boxes):
            conf = int(data['conf'][j])
            text = data['text'][j].strip()
            
            if conf > 30 and len(text) > 0:
                x, y, w, h = data['left'][j], data['top'][j], data['width'][j], data['height'][j]
                
                key = (data['block_num'][j], data['par_num'][j])
                if key not in paragraphs:
                    paragraphs[key] = {
                        'text_list': [], 
                        'rects': [], 
                        'heights': [], 
                        'line_nums': set()
                    }
                
                paragraphs[key]['text_list'].append(text)
                paragraphs[key]['rects'].append((x, y, w, h))
                paragraphs[key]['heights'].append(h)
                paragraphs[key]['line_nums'].add(data['line_num'][j])

        # --- 第二階段：計算最大字體 (找標題) ---
        max_font_size_on_page = 0
        for key in paragraphs:
            f_size = get_font_size_float(paragraphs[key]['heights'])
            paragraphs[key]['calculated_size'] = f_size
            if f_size > max_font_size_on_page:
                max_font_size_on_page = f_size

        # --- 第三階段：智慧過濾邏輯 ---
        for key, p_data in paragraphs.items():
            full_text = " ".join(p_data['text_list'])
            all_rects = p_data['rects']
            
            min_x = min([r[0] for r in all_rects])
            min_y = min([r[1] for r in all_rects])
            max_x2 = max([r[0] + r[2] for r in all_rects])
            max_y2 = max([r[1] + r[3] for r in all_rects])
            p_w = max_x2 - min_x
            p_h = max_y2 - min_y
            
            # 1.【NotebookLM 移除】
            if "notebook" in full_text.lower() and min_y > (img_h * 0.8):
                bg_color = get_smart_median_color(img_np, min_x, min_y, p_w, p_h)
                cv2.rectangle(clean_bg_img, (min_x-2, min_y-2), (max_x2+2, max_y2+2), bg_color, -1)
                continue 

            # 2.【屬性判斷】
            is_bullet = is_list_item(full_text)  # 是否為清單
            is_black = is_text_black(img_np, min_x, min_y, p_w, p_h) # 是否為黑色
            is_title = (p_data['calculated_size'] >= max_font_size_on_page - 2) and (max_font_size_on_page > 14) # 是否為標題
            is_multiline = len(p_data['line_nums']) >= 2 # 是否多行
            
            # 3.【拆解決策樹】
            should_extract = False
            
            if is_bullet:
                # 規則 A: 如果是清單 -> 不論顏色、不論行數，強制拆解！
                should_extract = True
            elif is_black:
                # 規則 B: 如果是黑色 -> 標題或多行才拆，單行不拆
                if is_title or is_multiline:
                    should_extract = True
            
            # 4.【執行動作】
            if should_extract:
                # 塗掉背景
                bg_color = get_smart_median_color(img_np, min_x, min_y, p_w, p_h)
                cv2.rectangle(clean_bg_img, (min_x-2, min_y-2), (max_x2+2, max_y2+2), bg_color, -1)
                
                # 標記要轉文字框
                p_data['should_export'] = True
                p_data['bbox'] = (min_x, min_y, p_w, p_h)
            else:
                # 保留在圖片上
                p_data['should_export'] = False

        # --- 第四階段：產生 PPT ---
        
        # 插入背景
        clean_bg_rgb = cv2.cvtColor(clean_bg_img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(clean_bg_rgb)
        img_stream = io.BytesIO()
        pil_img.save(img_stream, format='JPEG', quality=95)
        img_stream.seek(0)
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        # 建立文字框
        scale_x = prs.slide_width / img_w
        scale_y = prs.slide_height / img_h
        
        for key, p_data in paragraphs.items():
            if not p_data.get('should_export'):
                continue
                
            min_x, min_y, p_w, p_h = p_data['bbox']
            full_text = " ".join(p_data['text_list'])
            
            ppt_x = min_x * scale_x
            ppt_y = min_y * scale_y
            ppt_w = p_w * scale_x + Inches(0.15)
            ppt_h = p_h * scale_y
            
            this_font_size = p_data['calculated_size']

            try:
                txBox = slide.shapes.add_textbox(ppt_x, ppt_y, ppt_w, ppt_h)
                tf = txBox.text_frame
                tf.word_wrap = True
                tf.text = full_text
                
                for paragraph in tf.paragraphs:
                    paragraph.font.size = Pt(this_font_size)
                    paragraph.font.name = "Arial"
                    paragraph.font.color.rgb = RGBColor(0, 0, 0)
                    
                    if (this_font_size >= max_font_size_on_page - 2) and (max_font_size_on_page > 14):
                        paragraph.font.bold = True
                    else:
                        paragraph.font.bold = False
            except:
                pass
        
        progress_bar.progress((i + 1) / total_pages)

    status_text.text("✅ 轉換完成！")
    ppt_output = io.BytesIO()
    prs.save(ppt_output)
    ppt_output.seek(0)
    return ppt_output

# --- 介面主入口 ---
uploaded_file = st.file_uploader("📂 請上傳 PDF 檔案", type=["pdf"])

if uploaded_file is not None:
    if st.button("🚀 開始轉換"):
        try:
            original_filename = uploaded_file.name
            file_root, _ = os.path.splitext(original_filename)
            new_filename = f"{file_root}_Fixed.pptx"

            ppt_file = process_pdf(uploaded_file)
            st.success(f"🎉 處理成功！")
            
            st.download_button(
                label=f"📥 下載 {new_filename}",
                data=ppt_file,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
            st.info("💡 提示：如果線上報錯，請檢查 requirements.txt 是否包含 opencv-python-headless。")
