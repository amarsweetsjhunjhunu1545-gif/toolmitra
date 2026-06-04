FROM python:3.10-slim

# Create a non-root user
RUN groupadd -r appuser && useradd -r -g appuser appuser

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-script-provider-python \
    qpdf \
    ghostscript \
    fonts-noto-core \
    fonts-noto-extra \
    fonts-lohit-deva \
    tesseract-ocr \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Set permissions for the non-root user
RUN mkdir -p /app/uploads /app/outputs && \
    chown -R appuser:appuser /app

# Switch to the non-root user
USER appuser

ENV PORT=10000
EXPOSE 10000

CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 1"]
