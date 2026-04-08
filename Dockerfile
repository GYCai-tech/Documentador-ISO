# Dockerfile — Contenedor para el Asistente ISO con Chainlit
# Chainlit escucha por defecto en el puerto 8000.

FROM python:3.11-slim

# Instala dependencias del sistema si las necesitas (ej: para python-docx)
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instala dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia el código fuente
COPY . .

# Puerto de Chainlit
EXPOSE 8000

HEALTHCHECK CMD curl --fail http://localhost:8000/healthz || exit 1

# Comando de arranque
# Pista: chainlit run app.py --host 0.0.0.0 --port 8000
CMD ["chainlit", "run", "app.py", "--host", "0.0.0.0", "--port", "8000"]
