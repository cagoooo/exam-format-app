FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV EXAM_FORMAT_DATA_DIR=/tmp/exam-format-app
ENV FLASK_DEBUG=0

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      fonts-noto-cjk \
      fonts-noto-cjk-extra \
      fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD exec gunicorn --bind :${PORT:-8080} --workers 1 --threads 4 --timeout 240 app:app
