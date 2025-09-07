FROM node:18-alpine

WORKDIR /app

# Install curl for health checks
RUN apk add --no-cache curl

# Copy package files
COPY package*.json ./

# Install dependencies (production)
ENV NODE_ENV=production
RUN npm ci --only=production

# Copy source code
COPY . .

# No build step required; running via tsx in production

# Expose and set port (match healthcheck)
ENV PORT=3001
EXPOSE 3001

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:3001/health || exit 1

# Start the server (tsx runs src/index.ts)
CMD ["npm", "start"]