# Microsoft Graph MCP Server Dockerfile - Optimized Multi-Stage Build
FROM node:18-alpine AS builder

# Set working directory
WORKDIR /app

# Copy package files first for better caching
COPY package*.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm ci

# Copy only necessary source files (exclude node_modules, .git, etc.)
COPY api/ ./api/
COPY server.js ./
COPY types.d.ts ./
COPY tsconfig.json ./
COPY tsconfig.worker.json ./

# Build the application
RUN npm run build

# Production stage
FROM node:18-alpine AS production

# Install curl for health checks
RUN apk add --no-cache curl

# Create non-root user first (before copying files)
RUN addgroup -g 1001 -S nodejs && \
    adduser -S mcp -u 1001

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install only production dependencies
RUN npm ci --only=production && npm cache clean --force

# Copy built application from builder stage
COPY --from=builder /app/dist ./dist
COPY --from=builder /app/server.js ./
COPY --from=builder /app/types.d.ts ./

# Change ownership of only the app directory (much faster)
RUN chown -R mcp:nodejs /app
USER mcp

# Expose port
EXPOSE 3001

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD ["curl", "-f", "http://localhost:3001/health"]

# Start the server
CMD ["npm", "start"]

# Start the server
CMD ["npm", "start"]
