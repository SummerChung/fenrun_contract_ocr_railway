import streamlit as st
import pandas as pd
import datetime
import re
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

st.set_page_config(page_title="分潤增補協議書 OCR 工具", layout="centered")
st.title("分潤增補協議書 OCR 自動辨識工具")
st.markdown("請上傳掃描版 PDF（最多 10 份），系統將使用 OCR 自動辨識欄位並產出 Excel。")

uploaded_files = st.file_uploader("上傳 PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) > 10:
        st.error("最多僅能上傳 10 份 PDF。")
    else:
        results = []
        progress = st.progress(0)
        for i, file in enumerate(uploaded_files):
            try:
                images = convert_from_bytes(file.read(), dpi=300)
                full_text = ""
                for img in images:
                    text = pytesseract.image_to_string(img, lang="chi_tra+eng")
                    full_text += text + "\n"

                match_email = re.search(r'[\w\.-]+@[\w\.-]+', full_text)
                match_price = re.search(r"單價[:：]?[\s]*NT\$?([0-9,]+)", full_text)
                match_total = re.search(r"合計[:：]?[\s]*NT\$?([0-9,]+)", full_text)
                match_count = re.search(r"聯網機.*?[數量台]*[:：]?[\s]*([一二三四五六七八九十百千壹貳參肆伍陸柒捌玖拾佰仟]+)", full_text)
                match_commission = re.search(r"分潤乙方.*?新台幣.*?([一二三四五六七八九十百千壹貳參肆伍陸柒捌玖拾佰仟]+)元", full_text)

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

                results.append({
                    "檔名": file.name,
                    "Email": match_email.group(0) if match_email else "",
                    "單價": match_price.group(1).replace(",", "") if match_price else "",
                    "合計": match_total.group(1).replace(",", "") if match_total else "",
                    "台數": chinese_to_arabic(match_count.group(1)) if match_count else "",
                    "分潤金額": chinese_to_arabic(match_commission.group(1)) if match_commission else ""
                })
            except Exception as e:
                st.error(f"{file.name} 分析時發生錯誤：{e}")
            progress.progress((i + 1) / len(uploaded_files))

        if results:
            df = pd.DataFrame(results)
            today = datetime.datetime.today().strftime("%Y%m%d")
            filename = f"分潤增補匯總_{today}.xlsx"
            df.to_excel(filename, index=False)
            st.success("辨識完成！請下載 Excel：")
            with open(filename, "rb") as f:
                st.download_button("下載 Excel", f, file_name=filename)
        else:
            st.warning("無法辨識任何欄位，請確認 PDF 是否為掃描合約且為中文格式。")
