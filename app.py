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
st.set_page_config(page_title="PDF 轉 PPT (旗艦版)", layout="wide")

st.title("📄 PDF 轉 PPT：智慧修補 + 排版還原")
st.markdown("""
**旗艦版功能：**
1. **色塊修補**：使用「區域中位數」吸色，背景修補最乾淨，無模糊痕跡。
2. **原字級還原**：精確計算像素與 PPT 點數轉換。
3. **智慧標題**：掃描整頁，僅將「字體最大」的標題設為粗體。
""")

# --- 參數設定 ---
OCR_LANG = 'chi_tra+eng'
TARGET_DPI = 300

# --- 核心功能函式 ---

def get_smart_median_color(image_np, x, y, w, h):
    """
    區域中位數吸色：
    吸取文字框周圍區域的中位數顏色，
    有效抵抗雜訊，抓出最準確的背景色。
    """
    img_h, img_w, _ = image_np.shape
    
    # 優先吸取文字左邊 10px 寬的區域
    sample_w = 10
    sample_h = min(h, 10)
    
    x1 = max(0, x - sample_w)
    x2 = x
    y1 = y
    y2 = min(img_h, y + sample_h)
    
    # 如果左邊沒空間，改吸上面
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
    """計算字體大小 (浮點數)"""
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
    
    status_text.text("正在將 PDF 轉為高解析圖片 (300 DPI)...")
    images = convert_from_bytes(bytes_data, dpi=TARGET_DPI)
    total_pages = len(images)
    
    for i, img in enumerate(images):
        status_text.text(f"🔄 正在處理第 {i+1} / {total_pages} 頁 (分析排版 -> 修補背景 -> 重建文字)...")
        
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
                
                # --- 步驟 A: 吸色與修補 ---
                bg_color = get_smart_median_color(img_np, x, y, w, h)
                
                # 擴張遮罩 (padding=3) 確保蓋住邊緣
                pad = 3
                cv2.rectangle(clean_bg_img, (x-pad, y-pad), (x+w+pad, y+h+pad), bg_color, -1)
                
                # --- 步驟 B: 收集資料 ---
                key = (data['block_num'][j], data['par_num'][j])
                if key not in paragraphs:
                    paragraphs[key] = {'text_list': [], 'rects': [], 'heights': []}
                
                paragraphs[key]['text_list'].append(text)
                paragraphs[key]['rects'].append((x, y, w, h))
                paragraphs[key]['heights'].append(h)
        
        # --- 步驟 C: 找出本頁最大字體 ---
        max_font_size_on_page = 0
        for key in paragraphs:
            f_size = get_font_size_float(paragraphs[key]['heights'])
            paragraphs[key]['calculated_size'] = f_size
            if f_size > max_font_size_on_page:
                max_font_size_on_page = f_size
        
        # 2. 插入修補後的背景
        clean_bg_rgb = cv2.cvtColor(clean_bg_img, cv2.COLOR_BGR2RGB)
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
                    
                    # 智慧加粗判定
                    if (this_font_size >= max_font_size_on_page - 2) and (max_font_size_on_page > 14):
                        paragraph.font.bold = True
                    else:
                        paragraph.font.bold = False
            except:
                pass
        
        progress_bar.progress((i + 1) / total_pages)

    status_text.text("✅ 轉換完成！準備下載...")
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
            st.success("🎉 處理成功！")
            st.download_button(
                label="📥 下載 PPTX",
                data=ppt_file,
                file_name="Converted_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"❌ 發生錯誤：{e}")
            st.info("💡 提示：請確認 packages.txt 內的 tesseract 依賴是否已正確安裝。")