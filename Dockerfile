FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-dejavu \
    fontconfig \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f -v

WORKDIR /app
COPY . .

EXPOSE 8080
CMD ["python", "server.py"]
