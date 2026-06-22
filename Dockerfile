FROM python:3.12-slim

# Install LibreOffice for PDF conversion
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements first for better caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Bind to Render's injected PORT
CMD gunicorn app:app --bind 0.0.0.0:$PORT