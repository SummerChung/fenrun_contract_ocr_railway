import streamlit as st
import pandas as pd
import datetime
import re
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps

st.set_page_config(page_title="分潤增補協議書 OCR 工具", layout="centered")
st.title("分潤增補協議書 OCR 自動辨識工具（修正版）")
st.markdown("📄 上傳掃描 PDF（最多 10 份），系統將進行圖像預處理並辨識欄位資訊，產出 Excel 報表。")

uploaded_files = st.file_uploader("⬆️ 上傳 PDF", type="pdf", accept_multiple_files=True)

def preprocess_image(img):
    img = img.convert("L")  # 灰階
    img = ImageOps.autocontrast(img)  # 自動對比
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
        st.error("最多僅能上傳 10 份 PDF。")
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

                email_match = re.search(r'([\w\.-]+@[\w\.-]+)', full_text)
                price_match = re.search(r"[單价價][：:NT\$ ]*([0-9,]{2,5})", full_text)
                total_match = re.search(r"[合总總]計[：:NT\$ ]*([0-9,]{2,6})", full_text)
                count_match = re.search(r"[數数量台]+[:： ]*([一二三四五六七八九十壹貳參肆伍陸柒捌玖拾佰仟]+)", full_text)
                commission_match = re.search(r"[分酬][潤利][金費額]?[：:新台幣 ]*([一二三四五六七八九十壹貳參肆伍陸柒捌玖拾佰仟]+)", full_text)

                results.append({
                    "檔名": file.name,
                    "Email": email_match.group(1) if email_match else "未辨識",
                    "單價": price_match.group(1).replace(",", "") if price_match else "未辨識",
                    "合計": total_match.group(1).replace(",", "") if total_match else "未辨識",
                    "台數": chinese_to_arabic(count_match.group(1)) if count_match else "未辨識",
                    "分潤金額": chinese_to_arabic(commission_match.group(1)) if commission_match else "未辨識"
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
            st.warning("⚠️ 未能成功辨識任何欄位，請確認掃描品質與合約格式。")
