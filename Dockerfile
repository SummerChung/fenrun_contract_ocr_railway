FROM python:3.9

WORKDIR /app
COPY . /app

# 安裝必要套件（含 easyocr）
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# 安裝 pdf2image 依賴（如需）
RUN apt-get update && apt-get install -y poppler-utils libgl1

CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
