FROM python:3.11-slim

# Install LibreOffice, fonts, and dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    fonts-open-sans \
    fonts-liberation \
    fonts-dejavu \
    fontconfig \
    wget \
    unzip \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Montserrat font (used in the certificate)
RUN mkdir -p /usr/share/fonts/truetype/montserrat \
    && wget -q "https://fonts.google.com/download?family=Montserrat" -O /tmp/Montserrat.zip \
    && unzip -q /tmp/Montserrat.zip -d /tmp/Montserrat \
    && find /tmp/Montserrat -name "*.ttf" -exec cp {} /usr/share/fonts/truetype/montserrat/ \; \
    && rm -rf /tmp/Montserrat.zip /tmp/Montserrat

# Install Microsoft core fonts (includes Tahoma, Arial, etc.)
RUN echo "deb http://deb.debian.org/debian bookworm contrib" >> /etc/apt/sources.list \
    && apt-get update \
    && echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections \
    && apt-get install -y ttf-mscorefonts-installer --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Refresh font cache so LibreOffice picks up all fonts
RUN fc-cache -f -v

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

ENV PORT=8080

CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 120 app:app
