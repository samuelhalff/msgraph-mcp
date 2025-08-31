# Microsoft Graph MCP Server Dockerfile - Ultra Optimized
FROM node:18-alpine AS builder

# Set working directory
WORKDIR /app

# Copy package files first for better caching
COPY package*.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm ci

# Copy only necessary source files
COPY api/ ./api/
COPY server.js ./
COPY types.d.ts ./
COPY tsconfig.json ./

# Build the application
RUN npm run build

# Production stage - Ultra minimal
FROM node:18-alpine AS production

# Install curl for health checks and create user in one step
RUN apk add --no-cache curl && \
    addgroup -g 1001 -S nodejs && \
    adduser -S mcp -u 1001

# Set working directory
WORKDIR /app

# Copy and install production dependencies in one layer
COPY package*.json ./
RUN npm ci --only=production && npm cache clean --force

# Copy only the essential runtime files and set ownership in one command
COPY --from=builder --chown=mcp:nodejs /app/dist ./dist
COPY --from=builder --chown=mcp:nodejs /app/server.js ./

# Create logs directory with proper ownership
RUN mkdir -p logs && chown -R mcp:nodejs logs

# Switch to non-root user
USER mcp

# Expose port
EXPOSE 3001

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD ["curl", "-f", "http://localhost:3001/health"]

# Start the server
CMD ["node", "server.js"]
