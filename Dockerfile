# syntax=docker/dockerfile:1
# Multi-stage build for smaller final image

ARG PYTHON_VERSION=3.13-slim
FROM python:${PYTHON_VERSION} AS base

# Prevent Python from writing .pyc files and buffering stdout
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Set workdir
WORKDIR /app

# System deps (add as needed)
RUN apt-get update && apt-get install -y --no-install-recommends \
        curl \
    && rm -rf /var/lib/apt/lists/*

# Install dependencies early to leverage Docker layer caching
COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt \
    && pip install --no-cache-dir gunicorn==21.2.0

# Copy application source
COPY app.py extract_employee_shifts.py wsgi.py ./
COPY templates ./templates

# Create empty uploads directory (mounted as volume in compose)
RUN mkdir -p uploads

# Create non-root user with build-time configurable UID/GID matching host (override via build args)
ARG APP_UID=1001
ARG APP_GID=100
RUN groupadd -g ${APP_GID} appgroup \
    && useradd -m -u ${APP_UID} -g appgroup appuser \
    && chown -R appuser:appgroup /app
USER appuser

# Expose port
EXPOSE 5000

# Default environment variables (can override in compose/production)
ENV PORT=5000 \
    FLASK_DEBUG=0 \
    WEB_CONCURRENCY=4

# Healthcheck (basic)
HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD curl -sf http://127.0.0.1:5000/ || exit 1

# Gunicorn command (4 workers auto-tuned) - adjust workers per CPU needs
# Using wsgi:app (see wsgi.py)
CMD ["bash", "-c", "gunicorn --bind 0.0.0.0:${PORT} --workers ${WEB_CONCURRENCY} wsgi:app"]
