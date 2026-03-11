FROM python:3.11-slim

# Install LibreOffice and fontconfig
RUN apt-get update && apt-get install -y \
    libreoffice \
    fontconfig \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

# Install the real Windows fonts
RUN mkdir -p /usr/share/fonts/windows && \
    cp fonts/*.TTF /usr/share/fonts/windows/ && \
    fc-cache -f -v

EXPOSE 8080
CMD ["python", "server.py"]
