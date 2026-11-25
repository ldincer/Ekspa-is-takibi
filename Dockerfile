FROM python:3.12-slim

# Loglar anında görünsün, saat/tarih Türkiye saatine göre olsun (SLA ve "bugün" hesapları için)
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=Europe/Istanbul \
    EKSPA_DATA_DIR=/veri

WORKDIR /app
COPY app.py ./
COPY static ./static
COPY data/baslangic_firsatlar.json data/baslangic_envanter.json data/baslangic_talepler.json ./data/

# Not: Railway Dockerfile'da VOLUME komutuna izin vermez. Kalıcı veri için:
#   Railway  -> servise Volume ekleyin, Mount path: /veri
#   Docker   -> docker run -v ekspa-veri:/veri ...
# Port: Railway PORT değişkenini verir, uygulama onu kullanır; yoksa 8080.
EXPOSE 8080
CMD ["python", "app.py"]
