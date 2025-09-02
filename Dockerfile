# Microsoft Graph MCP Server Dockerfile - Optimized
FROM node:18-alpine AS builder

WORKDIR /app

# Copy package files first for better caching
COPY package*.json ./

# Install dependencies (including dev for build)
RUN npm ci

# Copy source files
COPY api/ ./api/
COPY types.d.ts ./
COPY tsconfig.json ./

# Build the app
RUN npm run build

# -------------------------
# Production stage
# -------------------------
FROM node:18-alpine AS production

RUN apk add --no-cache curl && \
    addgroup -g 1001 -S nodejs && \
    adduser -S mcp -u 1001

WORKDIR /app

# Copy and install only prod dependencies
COPY package*.json ./
RUN npm ci --only=production && npm cache clean --force

# Copy compiled output
COPY --from=builder --chown=mcp:nodejs /app/dist ./dist

# Switch to non-root user
USER mcp

EXPOSE 3001

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD ["curl", "-f", "http://localhost:3001/health"]

# Start server
CMD ["node", "dist/api/index.js"]
