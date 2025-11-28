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
import re

# --- 網頁設定 ---
st.set_page_config(page_title="PDF 轉 PPT (圖片避讓版)", layout="wide")

st.title("📄 PDF 轉 PPT：圖片避讓 + 智慧過濾")
st.markdown("""
**本次更新邏輯：**
1. **圖片避讓**：自動偵測頁面上的「大圖片/圖表」，凡是 **壓在圖上** 或 **緊鄰圖片** 的文字，一律保留在背景不拆解。
2. **清單強化**：條列式清單 (`•`, `1.`) 強制拆解。
3. **干擾移除**：NotebookLM 浮水印移除。
""")

# --- 參數設定 ---
OCR_LANG = 'chi_tra+eng'
TARGET_DPI = 300
BLACK_THRESHOLD = 80 

# --- 核心功能 ---

def get_large_image_mask(image_np, text_boxes):
    """
    產生「圖片禁區遮罩」。
    邏輯：
    1. 把原圖二值化。
    2. 把所有「文字位置」塗白 (消除文字干擾)。
    3. 剩下的就是「圖形/線條/照片」。
    4. 找出這些圖形的輪廓，過濾掉太小的雜訊。
    5. 將大圖形的位置標記出來，並往外擴張 (膨脹)，形成禁區。
    """
    img_h, img_w, _ = image_np.shape
    
    # 1. 轉灰階並二值化 (黑底白線)
    gray = cv2.cvtColor(image_np, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # 2. 把偵測到的「文字」全部塗黑 (在二值圖中，背景是黑，前景是白，所以我們要塗黑文字讓它消失)
    # 修正：binary 是黑底白前，所以要把文字區域塗黑(0)
    for (tx, ty, tw, th) in text_boxes:
        # 稍微擴大一點塗抹，確保文字徹底消失
        cv2.rectangle(binary, (max(0, tx-5), max(0, ty-5)), (tx+tw+5, ty+th+5), 0, -1)
        
    # 3. 膨脹處理，讓破碎的圖形線條連在一起
    kernel = np.ones((5,5), np.uint8)
    dilated = cv2.dilate(binary, kernel, iterations=2)
    
    # 4. 找輪廓 (這些就是圖片/圖表)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 建立禁區遮罩 (白底黑字概念，這裡用 255 代表禁區)
    danger_zone_mask = np.zeros((img_h, img_w), dtype=np.uint8)
    
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        
        # 條件：面積夠大才算「大圖片」 (例如頁面面積的 2% 以上)
        # 避免把小 icon 或分隔線當成大圖
        if area > (img_w * img_h * 0.02):
            cv2.rectangle(danger_zone_mask, (x, y), (x+w, y+h), 255, -1)
            
    # 5. 將禁區再往外擴張一點 (Buffer)，讓靠近圖片的字也受到保護
    buffer_kernel = np.ones((15, 15), np.uint8) # 擴張約 7px
    danger_zone_mask = cv2.dilate(danger_zone_mask, buffer_kernel, iterations=1)
    
    return danger_zone_mask

def is_touching_image(x, y, w, h, danger_mask):
    """檢查文字框是否撞到圖片禁區"""
    # 取出文字框在 mask 對應的區域
    roi = danger_mask[y:y+h, x:x+w]
    # 如果區域內有任何白色像素 (255)，代表撞到了
    return cv2.countNonZero(roi) > 0

def is_list_item(text):
    """判斷是否為清單"""
    text = text.strip()
    if not text: return False
    markers = ['•', '●', '○', '▪', '▫', '◆', '◇', '➢', '➣', '➤', '→', '-', '—', '–', '*', '>']
    if any(text.startswith(m) for m in markers): return True
    pattern = r'^(\d+|[a-zA-Z])[\.\)]\s+'
    if re.match(pattern, text): return True
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
    x1 = max(0, x - sample_w); x2 = x
    y1 = y; y2 = min(img_h, y + min(h, 10))
    if (x2 - x1) < 2:
        x1 = x; x2 = min(img_w, x + sample_w)
        y1 = max(0, y - 5); y2 = y
    try:
        roi = image_np[y1:y2, x1:x2]
        if roi.size == 0: return (255, 255, 255)
        median_color = np.median(roi, axis=(0, 1))
        return (int(median_color[0]), int(median_color[1]), int(median_color[2]))
    except:
        return (255, 255, 255)

def get_font_size_float(heights_px):
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
        all_text_boxes = [] # 用來存所有文字位置，給圖片偵測用
        n_boxes = len(data['text'])
        
        clean_bg_img = img_np.copy()
        
        # --- 第一階段：收集資料 ---
        for j in range(n_boxes):
            conf = int(data['conf'][j])
            text = data['text'][j].strip()
            
            if conf > 30 and len(text) > 0:
                x, y, w, h = data['left'][j], data['
