# 1. Use Node.js base image
FROM node:18-slim

# 2. Install required tools (LibreOffice, Ghostscript, QPDF, Python, Poppler)
RUN apt-get update && apt-get install -y \
    libreoffice \
    ghostscript \
    qpdf \
    poppler-utils \
    python3 \
    python3-pip \
    chromium \
    && rm -rf /var/lib/apt/lists/*

# 3. Set Puppeteer to use installed Chromium
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

# 4. Create App Directory
WORKDIR /app

# 5. Copy package files and install dependencies
COPY package*.json ./
RUN npm install

# 6. Copy python requirements (if any) or install pdf2docx directly
RUN pip3 install pdf2docx --break-system-packages

# 7. Copy the rest of the code
COPY . .

# 8. Expose Port
EXPOSE 3000

# 9. Start Server
CMD ["node", "server.js"]