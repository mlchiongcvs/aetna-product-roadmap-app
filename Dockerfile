FROM node:20-alpine

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY dist/ ./dist/
COPY server/ ./server/

EXPOSE 8080

CMD ["node", "server/index.js"]
