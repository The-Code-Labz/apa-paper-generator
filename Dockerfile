FROM python:3.12-slim

LABEL description="APA 7th Edition Paper Generator — FastAPI + python-docx"
LABEL org.opencontainers.image.source="https://github.com/The-Code-Labz/apa-paper-generator"
LABEL org.opencontainers.image.authors="The-Code-Labz"

# Install curl for the healthcheck
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

# ── Non-root user ──────────────────────────────────────────────────────────────
# Running as root inside containers is unnecessary attack surface.
RUN groupadd --system apa \
    && useradd --system --gid apa --no-create-home --shell /sbin/nologin apa

WORKDIR /app

# Install Python dependencies first (layer caching — code changes won't bust this)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and static assets
COPY app.py .
COPY static/ ./static/

# Pre-create the output directory with correct ownership
# The compose volume mounts here; the directory must exist before the volume binds.
RUN mkdir -p /tmp/apa-papers && chown apa:apa /tmp/apa-papers

USER apa

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -sf http://localhost:8000/api/health || exit 1

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
