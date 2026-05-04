FROM node:20-alpine

WORKDIR /app

COPY package.json ./
RUN npm install --omit=dev

COPY server.js ./
COPY pds-generator.js ./
COPY public/ ./public/
COPY templates/ ./templates/

EXPOSE 3000

CMD ["node", "server.js"]
