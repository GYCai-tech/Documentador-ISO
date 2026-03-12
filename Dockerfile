# ── Etapa base: Python + Node.js (para mermaid-cli) ──────────────────────────
FROM python:3.11-slim

# Dependencias del sistema: Node.js, Chromium (headless para mmdc/puppeteer)
RUN apt-get update && apt-get install -y --no-install-recommends \
    nodejs \
    npm \
    chromium \
    fonts-liberation \
    libgbm1 \
    && rm -rf /var/lib/apt/lists/*

# mermaid-cli global
RUN npm install -g @mermaid-js/mermaid-cli

# Apuntar puppeteer al Chromium del sistema (evita descarga propia)
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

WORKDIR /app

# Instalar dependencias Python (capa cacheada separada)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código fuente
COPY . .

# Streamlit escucha en 8501
EXPOSE 8501

HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

CMD ["streamlit", "run", "app.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--browser.gatherUsageStats=false"]
