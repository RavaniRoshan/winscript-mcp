# WinScript MCP Server - Docker Image
# Windows automation server for AI agents
# 
# Build: docker build -t winscript:latest .
# Run:   docker run -v %USERPROFILE%/.winscript:~/.winscript winscript:latest
#
# Note: This Docker image works on Linux hosts but targets Windows containers.
# For production Windows deployment, use Windows Server Core or nanoserver base.

FROM python:3.11-slim

LABEL maintainer="Roshan Ravani"
LABEL description="WinScript MCP Server - AppleScript for Windows"
LABEL version="0.1.0"

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create directory for persistent data
RUN mkdir -p /root/.winscript/workflows

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1

# Expose port for MCP communication (if using HTTP transport in future)
EXPOSE 8080

# Volume for persistent data (audit logs, memory, workflows)
VOLUME ["/root/.winscript"]

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "from winscript import mcp; print('WinScript MCP server healthy')" || exit 1

# Run the MCP server
CMD ["python", "-m", "winscript.server"]
