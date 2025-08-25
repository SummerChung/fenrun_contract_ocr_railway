import streamlit as st
import pandas as pd
import datetime
import re
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps

st.set_page_config(page_title="分潤增補 OCR 工具（進階強化版）", layout="centered")
st.title("📄 分潤增補協議書 OCR 自動辨識工具（進階強化版）")
st.markdown("針對表格型 PDF、手寫 Email、國字金額進行強化處理。支援最多 10 份掃描 PDF。")

uploaded_files = st.file_uploader("⬆️ 上傳 PDF 檔案", type="pdf", accept_multiple_files=True)

def preprocess_image(img):
    img = img.convert("L")  # 灰階
    img = ImageOps.autocontrast(img)
    img = img.point(lambda x: 0 if x < 180 else 255, '1')  # 二值化
    return img

def chinese_to_arabic(chinese_num):
    map = {'零':0,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,
           '百':100,'千':1000,'萬':10000,'壹':1,'貳':2,'參':3,'肆':4,'伍':5,'陸':6,'柒':7,'捌':8,'玖':9,'拾':10,'佰':100,'仟':1000}
    unit = 0
    ldig = []
    for c in reversed(chinese_num):
        val = map.get(c, None)
        if val is None:
            continue
        if val >= 10:
            if val > unit:
                unit = val
                ldig.append(unit)
            else:
                ldig[-1] += val
        else:
            if unit:
                ldig.append(val * unit)
            else:
                ldig.append(val)
    return sum(ldig) if ldig else 0

if uploaded_files:
    if len(uploaded_files) > 10:
        st.error("最多可上傳 10 份 PDF。")
    else:
        results = []
        progress = st.progress(0)
        for i, file in enumerate(uploaded_files):
            st.write(f"📄 處理中：{file.name}")
            try:
                images = convert_from_bytes(file.read(), dpi=300)
                full_text = ""
                for img in images:
                    processed = preprocess_image(img)
                    ocr_text = pytesseract.image_to_string(processed, lang="chi_tra+eng")
                    full_text += ocr_text + "\n"

                # Email：標準與模糊格式
                email_match = re.search(r'([\w\.-]+@[\w\.-]+)', full_text)
                email_fuzzy = re.search(r'([\w\(\)\.-]+\s*@\s*[\w\.-]+)', full_text)

                email_result = "未辨識"
                if email_match:
                    email_result = email_match.group(1)
                elif email_fuzzy:
                    email_result = "疑似：" + email_fuzzy.group(1).replace(" ", "")

                # 起始年月
                start_match = re.search(r"(\d{3})年(\d{1,2})月", full_text)
                start_date = f"{start_match.group(1)}年{start_match.group(2)}月" if start_match else "未辨識"

                # 台數
                count_match = re.search(r"EIP聯網機\s*([壹貳參肆伍陸柒捌玖拾佰仟一二三四五六七八九十]+)台", full_text)
                tai_result = chinese_to_arabic(count_match.group(1)) if count_match else "未辨識"

                # 單價與合計
                price_match = re.search(r"NT\$\s*([0-9,]+)", full_text)
                total_match = re.search(r"合計\s*NT\$?\s*([0-9,]+)", full_text)

                price_result = price_match.group(1).replace(",", "") if price_match else "未辨識"
                total_result = total_match.group(1).replace(",", "") if total_match else "未辨識"

                # 分潤金額
                commission_match = re.search(r"分潤乙方.*?新[台臺]幣[\s]*([壹貳參肆伍陸柒捌玖拾佰仟一二三四五六七八九十百千]+)元", full_text)
                commission_result = chinese_to_arabic(commission_match.group(1)) if commission_match else "未辨識"

                results.append({
                    "檔名": file.name,
                    "起始年月": start_date,
                    "Email": email_result,
                    "台數": tai_result,
                    "單價": price_result,
                    "合計": total_result,
                    "分潤金額": commission_result
                })
            except Exception as e:
                st.error(f"❌ {file.name} 發生錯誤：{e}")
            progress.progress((i + 1) / len(uploaded_files))

        if results:
            df = pd.DataFrame(results)
            today = datetime.datetime.today().strftime("%Y%m%d")
            filename = f"分潤增補匯總_{today}.xlsx"
            st.success("✅ 辨識完成！請下載 Excel：")
            with open(filename, "wb") as f:
                df.to_excel(f, index=False)
            with open(filename, "rb") as f:
                st.download_button("⬇️ 下載 Excel", f, file_name=filename)
        else:
            st.warning("⚠️ 所有檔案均無辨識結果，請確認掃描品質或聯繫開發人員。")
