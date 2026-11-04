FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-impress && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py merge.py ./
COPY templates ./templates
COPY static ./static
RUN useradd -r -u 10001 hub && mkdir -p /data && chown hub:hub /data
USER hub
ENV PPT_HUB_DATA_DIR=/data PPT_HUB_HOST=0.0.0.0 PPT_HUB_PORT=1410
EXPOSE 1410
CMD ["python", "app.py"]
