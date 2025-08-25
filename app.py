import streamlit as st
import pandas as pd
import datetime
import re
import easyocr
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps

st.set_page_config(page_title="分潤增補 OCR 工具（EasyOCR 版）", layout="centered")
st.title("📄 分潤增補協議書 OCR 工具（EasyOCR 版）")
st.markdown("使用 EasyOCR 辨識圖像 PDF，可更好處理手寫或印刷細字。")

uploaded_files = st.file_uploader("⬆️ 上傳 PDF 檔案", type="pdf", accept_multiple_files=True)

reader = easyocr.Reader(['ch_tra', 'en'], gpu=False)

def preprocess_image(img):
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
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

def extract_fields(text):
    start_match = re.search(r"(\d{3})年(\d{1,2})月", text)
    start_date = f"{start_match.group(1)}年{start_match.group(2)}月" if start_match else "未辨識"

    count_match = re.search(r"EIP聯網機[\s:]*([壹貳參肆伍陸柒捌玖拾佰仟一二三四五六七八九十]+)台", text)
    count = chinese_to_arabic(count_match.group(1)) if count_match else "未辨識"

    price_match = re.search(r"單價[\sNT$]*([0-9,]+)", text)
    price = price_match.group(1).replace(",", "") if price_match else "未辨識"

    total_match = re.search(r"合計[\sNT$]*([0-9,]+)", text)
    total = total_match.group(1).replace(",", "") if total_match else "未辨識"

    commission_match = re.search(r"分潤乙方.*?幣[\s]*([壹貳參肆伍陸柒捌玖拾佰仟一二三四五六七八九十百千]+)元", text)
    commission = chinese_to_arabic(commission_match.group(1)) if commission_match else "未辨識"

    email_match = re.search(r'([\w\.-]+@[\w\.-]+)', text)
    email = email_match.group(1) if email_match else "未辨識"

    return start_date, count, price, total, commission, email

if uploaded_files:
    results = []
    progress = st.progress(0)
    for i, file in enumerate(uploaded_files):
        st.write(f"📄 處理中：{file.name}")
        try:
            images = convert_from_bytes(file.read(), dpi=300)
            full_text = ""
            for img in images:
                processed = preprocess_image(img)
                result = reader.readtext(np.array(processed), detail=0, paragraph=True)
                full_text += "\n".join(result)

            start, count, price, total, commission, email = extract_fields(full_text)
            results.append({
                "檔名": file.name,
                "起始年月": start,
                "Email": email,
                "台數": count,
                "單價": price,
                "合計": total,
                "分潤金額": commission
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
