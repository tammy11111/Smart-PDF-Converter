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

# --- 網頁設定 ---
st.set_page_config(page_title="PDF 轉 PPT (強力去字版)", layout="wide")

st.title("📄 PDF 轉 PPT：強力去字 + 智慧排版")
st.markdown("""
**本次更新重點：**
1. **全域遮罩膨脹 (Mask Dilation)**：自動將文字選取範圍「外擴」，確保 g, y, j 等字母尾巴完全清除。
2. **一次性修補**：避免重複塗抹造成的背景髒污。
""")

# --- 參數設定 ---
OCR_LANG = 'chi_tra+eng'
TARGET_DPI = 300

# --- 核心功能 ---

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
        
        # 準備影像 (OpenCV BGR)
        img_np = np.array(img)
        img_np = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        img_h, img_w, _ = img_np.shape
        
        # 1. 執行 OCR
        data = pytesseract.image_to_data(img, lang=OCR_LANG, output_type=Output.DICT)
        
        paragraphs = {}
        n_boxes = len(data['text'])
        
        # 建立一個「全頁遮罩」 (一開始全黑)
        full_mask = np.zeros(img_np.shape[:2], dtype=np.uint8)
        
        # --- 第一階段：標記所有文字位置 ---
        for j in range(n_boxes):
            conf = int(data['conf'][j])
            text = data['text'][j].strip()
            
            if conf > 30 and len(text) > 0:
                x, y, w, h = data['left'][j], data['top'][j], data['width'][j], data['height'][j]
                
                # 在遮罩上畫白色矩形 (標記這裡是文字)
                cv2.rectangle(full_mask, (x, y), (x+w, y+h), 255, -1)
                
                # 收集資料供後續 PPT 使用
                key = (data['block_num'][j], data['par_num'][j])
                if key not in paragraphs:
                    paragraphs[key] = {'text_list': [], 'rects': [], 'heights': []}
                
                paragraphs[key]['text_list'].append(text)
                paragraphs[key]['rects'].append((x, y, w, h))
                paragraphs[key]['heights'].append(h)
        
        # --- 第二階段：遮罩膨脹 (Dilation) - 關鍵步驟！ ---
        # 這一步會把剛剛畫的所有白框「變胖」，確保蓋住文字邊緣的殘影
        # kernel 設為 3x3，膨脹 2 次，相當於往外擴張約 4-6 像素
        kernel = np.ones((3, 3), np.uint8)
        dilated_mask = cv2.dilate(full_mask, kernel, iterations=2)
        
        # --- 第三階段：一次性背景修補 ---
        # 使用 Telea 演算法，根據膨脹後的遮罩進行修補
        if np.sum(dilated_mask) > 0:
            # radius=5 (參考周圍 5px 的顏色來補)
            inpainted_img = cv2.inpaint(img_np, dilated_mask, 5, cv2.INPAINT_TELEA)
        else:
            inpainted_img = img_np

        # --- 第四階段：計算最大字體 (智慧標題) ---
        max_font_size_on_page = 0
        for key in paragraphs:
            f_size = get_font_size_float(paragraphs[key]['heights'])
            paragraphs[key]['calculated_size'] = f_size
            if f_size > max_font_size_on_page:
                max_font_size_on_page = f_size
        
        # 2. 插入修補後的背景
        clean_bg_rgb = cv2.cvtColor(inpainted_img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(clean_bg_rgb)
        img_stream = io.BytesIO()
        pil_img.save(img_stream, format='JPEG', quality=95)
        img_stream.seek(0)
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        # 3. 貼上文字框
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
            ppt_file = process_pdf(uploaded_file)
            st.success("🎉 處理成功！背景已強力清除。")
            st.download_button(
                label="📥 下載 PPTX",
                data=ppt_file,
                file_name="Clean_Fixed.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
            st.info("💡 如果出現 cv2 錯誤，請確認 requirements.txt 包含 opencv-python-headless。")
