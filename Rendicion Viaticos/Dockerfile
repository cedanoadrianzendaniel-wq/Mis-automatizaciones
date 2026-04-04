FROM node:18-slim

WORKDIR /app

# Instalar dependencias del sistema para Tesseract OCR
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

COPY package*.json ./
RUN npm install --production

COPY . .

RUN mkdir -p uploads

EXPOSE 3001

CMD ["node", "server.js"]
