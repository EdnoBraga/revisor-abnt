FROM python:3.12-slim

# Permite gravar o commit git usado no build, exposto depois em /api/health,
# para nunca mais ficar em duvida sobre qual versao do motor esta rodando.
ARG GIT_COMMIT=unknown

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    REVISOR_ABNT_JOBS_DIR=/var/lib/revisor-abnt/jobs \
    REVISOR_ABNT_COMMIT=${GIT_COMMIT} \
    HOME=/tmp

RUN apt-get update \
    && apt-get install --no-install-recommends -y libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app
COPY revisor-abnt-docx ./revisor-abnt-docx
RUN mkdir -p /var/lib/revisor-abnt/jobs
RUN addgroup --system app && adduser --system --ingroup app app \
    && chown -R app:app /app /var/lib/revisor-abnt

VOLUME ["/var/lib/revisor-abnt"]
USER app

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 CMD python -c "from urllib.request import urlopen; urlopen('http://127.0.0.1:8000/api/health', timeout=3)"
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
