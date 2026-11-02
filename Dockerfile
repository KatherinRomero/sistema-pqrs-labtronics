FROM ghcr.io/puppeteer/puppeteer:23
USER root
WORKDIR /app
ENV PUPPETEER_SKIP_DOWNLOAD=true
COPY package.json ./
RUN npm install --omit=dev
COPY . .
RUN mkdir -p data && chown -R pptruser:pptruser /app
USER pptruser
CMD ["node","server.js"]
