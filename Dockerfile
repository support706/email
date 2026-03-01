FROM python:3.11-slim

# Install LibreOffice and dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Verify soffice is available
RUN soffice --version

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

ENV PORT=8080

CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 120 app:app
