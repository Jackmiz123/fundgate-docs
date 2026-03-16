FROM python:3.11-slim

# Install LibreOffice and font tools
RUN apt-get update && apt-get install -y \
    libreoffice \
    fontconfig \
    fonts-liberation \
    fonts-dejavu \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir python-docx

# Install the real Windows fonts
RUN mkdir -p /usr/share/fonts/windows && \
    find /app/fonts -name "*.TTF" -o -name "*.ttf" | xargs -I{} cp {} /usr/share/fonts/windows/ && \
    fc-cache -f -v

EXPOSE 8080
CMD ["python", "server.py"]
