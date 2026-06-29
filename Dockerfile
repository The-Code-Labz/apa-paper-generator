FROM python:3.12-slim

LABEL description="APA 7th Edition Paper Generator — FastAPI + python-docx"

# Install curl (used by the healthcheck)
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies (separate layer so rebuilds are fast)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and static assets
COPY app.py .
COPY static/ ./static/

# Pre-create the output directory; the compose volume will mount here
RUN mkdir -p /tmp/apa-papers

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
