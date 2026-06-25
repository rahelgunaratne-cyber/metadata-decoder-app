# Metadata Decoder — single-image build for Cloud Run.
#
# Stage 1 builds the React SPA. Stage 2 is the Python runtime that serves the
# JSON API and the built SPA from one process, so the whole product is one
# Cloud Run service on one URL.

# ---- Stage 1: build the React SPA -----------------------------------------
FROM node:20-alpine AS frontend
WORKDIR /app/frontend

# Install deps first so this layer is cached unless package files change.
COPY frontend/package*.json ./
RUN npm ci

COPY frontend/ ./
RUN npm run build


# ---- Stage 2: Python runtime (API + static SPA) ---------------------------
FROM python:3.11-slim AS runtime

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    ENV=prod \
    PORT=8080

WORKDIR /app

COPY backend/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Backend source (app/ package + vendored engine/).
COPY backend/ ./

# Built SPA -> FastAPI serves it from /app/static (see app/config.py).
COPY --from=frontend /app/frontend/dist ./static

EXPOSE 8080

# Cloud Run injects $PORT; default to 8080 for local `docker run`.
CMD ["sh", "-c", "exec uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8080}"]
