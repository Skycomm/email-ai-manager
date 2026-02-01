# Email AI Manager
# Multi-stage build for smaller final image

FROM python:3.11-slim as builder

WORKDIR /build

# Install build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Create virtual environment
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt


# Final stage
FROM python:3.11-slim

WORKDIR /app

# Create non-root user for security
RUN groupadd -r emailai && useradd -r -g emailai emailai

# Copy virtual environment from builder
COPY --from=builder /opt/venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Copy application code
COPY app/ ./app/
COPY migrations/ ./migrations/
COPY dashboard/ ./dashboard/

# Create data directory
RUN mkdir -p /data && chown -R emailai:emailai /data

# Switch to non-root user
USER emailai

# Environment defaults (override in docker-compose or runtime)
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    LOG_LEVEL=INFO \
    DB_PATH=/data/email_manager.db \
    POLL_INTERVAL_SECONDS=60

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import sqlite3; sqlite3.connect('$DB_PATH').execute('SELECT 1')" || exit 1

# Volume for persistent data
VOLUME ["/data"]

# Expose dashboard port
EXPOSE 8080

# Run the application
CMD ["python", "-m", "app.main"]
