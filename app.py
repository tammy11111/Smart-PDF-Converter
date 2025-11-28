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

# --- 網頁設定 ---
st.set_page_config(page_title="PDF 轉 PPT (顏色過濾版)", layout="wide")

st.title("📄 PDF 轉 PPT：智慧顏色過濾 + 色塊修補")
st.markdown("""
**本次更新重點：**
1. **顏色過濾**：只有 **「黑色/深灰色」** 的文字會被拆解成可編輯文字。
2. **保留圖解**：圖片中有顏色的文字（紅/藍/綠等）將自動保留在背景圖上，不會被破壞。
3. **背景修補**：黑色文字部分依然使用「智慧色塊」蓋除。
""")

# --- 參數設定 ---
OCR_LANG = 'chi_tra+eng'
TARGET_DPI = 300
# 定義「黑色」的門檻 (RGB 0~255)，數值越小越嚴格(越黑)
# 設定 80 允許深灰色也被視為內文
BLACK_THRESHOLD = 80 

# --- 核心功能 ---

def is_text_black(image_np, x, y, w, h):
    """
    判斷該區域的文字是否為黑色/深色。
    原理：
    1. 切出文字區域。
    2. 轉灰階並二值化，找出「文字像素」(前景)。
    3. 計算這些像素在原圖(RGB)中的平均顏色。
    4. 如果 R, G, B 都小於門檻，認定為黑色文字。
    """
    # 邊界檢查
    img_h, img_w, _ = image_np.shape
    x = max(0, x); y = max(0, y)
    w = min(w, img_w - x); h = min(h, img_h - y)
    
    if w <= 0 or h <= 0: return False

    roi = image_np[y:y+h, x:x+w]
    
    # 轉灰階
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    
    # 使用 Otsu 二值化找出文字像素 (黑色部分)
    # THRESH_BINARY_INV: 讓文字變白(255)，背景變黑(0)，方便做遮罩
    _, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # 如果找不到文字像素 (可能是全白)，直接回傳 False
    if cv2.countNonZero(mask) == 0:
        return False

    # 計算遮罩區域內的平均顏色 (B, G, R)
    mean_val = cv2.mean(roi, mask=mask)
    b, g, r = mean_val[0], mean_val[1], mean_val[2]
    
    # 判斷是否夠黑 (R, G, B 都必須很低)
    if b < BLACK_THRESHOLD and g < BLACK_THRESHOLD and r < BLACK_THRESHOLD:
        return True # 是黑色文字 -> 拆！
    else:
        return False # 是彩色文字 -> 不拆！

def get_smart_median_color(image_np, x, y, w, h):
    """區域中位數吸色"""
    img_h, img_w, _ = image_np.shape
    sample_w = 10
    sample_h = min(h, 10)
    
    x1 = max(0, x - sample_w)
    x2 = x
    y1 = y
    y2 = min(img_h, y + sample_h)
    
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
        status_text.text(f"🔄 正在處理第 {i+1} / {total_pages} 頁 (正在過濾彩色文字)...")
        
        # 準備影像
        img_np = np.array(img)
        img_np = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        img_h, img_w, _ = img_np.shape
        
        # 1. 執行 OCR
        data = pytesseract.image_to_data(img, lang=OCR_LANG, output_type=Output.DICT)
        
        paragraphs = {}
        n_boxes = len(data['text'])
        
        # 複製背景圖來修補
        clean_bg_img = img_np.copy()
        
        for j in range(n_boxes):
            conf = int(data['conf'][j])
            text = data['text'][j].strip()
            
            if conf > 30 and len(text) > 0:
                x, y, w, h = data['left'][j], data['top'][j], data['width'][j], data['height'][j]
                
                # --- 關鍵判斷：是黑色文字嗎？ ---
                if is_text_black(img_np, x, y, w, h):
                    # 【情況 A：黑色/深色文字】-> 執行「拆解」SOP
                    
                    # 1. 吸取背景色
                    bg_color = get_smart_median_color(img_np, x, y, w, h)
                    
                    # 2. 塗掉背景 (pad=2)
                    pad = 2
                    cv2.rectangle(clean_bg_img, (x-pad, y-pad), (x+w+pad, y+h+pad), bg_color, -1)
                    
                    # 3. 收集資料準備轉文字框
                    key = (data['block_num'][j], data['par_num'][j])
                    if key not in paragraphs:
                        paragraphs[key] = {'text_list': [], 'rects': [], 'heights': []}
                    
                    paragraphs[key]['text_list'].append(text)
                    paragraphs[key]['rects'].append((x, y, w, h))
                    paragraphs[key]['heights'].append(h)
                
                else:
                    # 【情況 B：彩色文字】-> 跳過
                    # 不塗背景，也不加入 paragraphs
                    # 這樣它就會留在原本的背景圖上
                    pass
        
        # 2. 計算頁面最大字體 (智慧標題用)
        max_font_size_on_page = 0
        for key in paragraphs:
            f_size = get_font_size_float(paragraphs[key]['heights'])
            paragraphs[key]['calculated_size'] = f_size
            if f_size > max_font_size_on_page:
                max_font_size_on_page = f_size
        
        # 3. 插入處理好的背景
        clean_bg_rgb = cv2.cvtColor(clean_bg_img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(clean_bg_rgb)
        img_stream = io.BytesIO()
        pil_img.save(img_stream, format='JPEG', quality=95)
        img_stream.seek(0)
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        # 4. 貼上文字框 (只會貼上被判定為黑色的文字)
        scale_x = prs.slide_width / img_w
        scale_y = prs.slide_height / img_h
        
        for key, p_data in paragraphs.items():
            full_text = " ".join(p_data['text_list'])
            all_rects = p_data['rects']
            
            min_x = min([r[0] for r in all_rects])
            min_y = min([r[1] for r in all_rects])
            max_x2 = max([r[0] + r[2] for r in all_rects])
            max_y2 = max([r[1] + r[3] for r in all_rects])
            
            ppt_x = min_x * scale_x
            ppt_y = min_y * scale_y
            ppt_w = (max_x2 - min_x) * scale_x + Inches(0.15)
            ppt_h = (max_y2 - min_y) * scale_y
            
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
            # 自動檔名處理
            original_filename = uploaded_file.name
            file_root, _ = os.path.splitext(original_filename)
            new_filename = f"{file_root}_Fixed.pptx"

            ppt_file = process_pdf(uploaded_file)
            st.success(f"🎉 處理成功！彩色文字已保留，黑色文字已轉換。")
            
            st.download_button(
                label=f"📥 下載 {new_filename}",
                data=ppt_file,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
            st.info("💡 提示：如果線上報錯，請檢查 requirements.txt 是否包含 opencv-python-headless。")
