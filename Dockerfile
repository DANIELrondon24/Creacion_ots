# Compatible with `podman build` and `docker build`
FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    TZ=America/Bogota \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    DATA_DIR=/app/data \
    LOG_DIR=/app/logs \
    HEADLESS=true

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
# Installs the Chromium matching the pip playwright version plus its system deps
RUN pip install -r requirements.txt \
    && python -m playwright install --with-deps chromium \
    && rm -rf /var/lib/apt/lists/*

COPY main.py server.py ./

RUN useradd --create-home --uid 1001 app \
    && mkdir -p /app/data /app/logs \
    && chown -R app:app /app
USER app

VOLUME ["/app/data", "/app/logs"]

CMD ["python", "server.py"]
