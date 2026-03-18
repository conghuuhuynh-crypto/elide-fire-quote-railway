FROM node:18-slim

# Cài LibreOffice
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libfontconfig1 \
      fonts-noto \
      fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json .
RUN npm install --production

COPY . .

# Tạo thư mục output
RUN mkdir -p outputs/quotes

EXPOSE 3333

CMD ["node", "server.js"]
