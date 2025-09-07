##############################
# Builder stage: compiles TS  #
##############################
FROM node:18-alpine AS builder

WORKDIR /app

# Install curl for optional health checks during build (kept minimal)
RUN apk add --no-cache curl

# Copy package files and install ALL deps (incl. dev) for tsc
COPY package*.json ./
RUN npm ci

# Copy source and build
COPY . .
RUN npm run build

################################
# Runtime stage: slim prod deps #
################################
FROM node:18-alpine AS runner

WORKDIR /app

# Install curl for health checks
RUN apk add --no-cache curl

# Install only production deps
COPY package*.json ./
RUN npm ci --omit=dev

# Copy compiled output from builder
COPY --from=builder /app/build ./build

# Expose port
EXPOSE 3001

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:3001/health || exit 1

# Start the server
CMD ["node", "build/index.js"]