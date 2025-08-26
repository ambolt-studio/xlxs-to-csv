# Dockerfile con fallback LibreOffice opcional
FROM node:18-slim

# Instala LibreOffice Calc (para evaluar fórmulas si activás USE_LIBREOFFICE=1)
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-calc fonts-liberation && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --production
COPY server.js ./

ENV NODE_ENV=production
ENV PORT=8080
# Ajustá si querés habilitar el fallback LO (0=off, 1=on)
ENV USE_LIBREOFFICE=1
# Payload máximo (MB) para JSON base64
# ENV MAX_PAYLOAD_MB=50

EXPOSE 8080
CMD ["node", "server.js"]
