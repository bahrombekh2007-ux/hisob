# Barqaror, wheel'lari tayyor bo'lgan Python versiyasi (3.14 muammosining oldini oladi)
FROM python:3.11-slim

# Loglar darhol chiqib tursin, .pyc fayllar yaratilmasin
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Avval faqat requirements'ni ko'chiramiz — Docker layer cache'dan foydalanish uchun
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Qolgan barcha loyiha fayllari
COPY . .

# Render PORT o'zgaruvchisini o'zi beradi (odatda 10000), lokal test uchun default
ENV PORT=5000
EXPOSE 5000

CMD ["python", "main.py"]
