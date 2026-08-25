FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    REVISOR_ABNT_JOBS_DIR=/var/lib/revisor-abnt/jobs

RUN apt-get update \
    && apt-get install --no-install-recommends -y libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app
COPY revisor-abnt-docx ./revisor-abnt-docx
RUN mkdir -p /var/lib/revisor-abnt/jobs

EXPOSE 8000
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
