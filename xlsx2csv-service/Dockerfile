# Dockerfile
FROM node:18-slim

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --production
COPY server.js ./

ENV NODE_ENV=production
ENV PORT=8080
# Optional: adjust payload max (MB)
# ENV MAX_PAYLOAD_MB=20

EXPOSE 8080
CMD ["node", "server.js"]
