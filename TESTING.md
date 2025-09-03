# Testing Guide

This document explains how to test the Microsoft Graph MCP server both locally and on your remote deployment.

## Test Script

The `test-server.js` script provides comprehensive testing of all server endpoints and functionality.

### Configuration

Before running tests, configure your environment variables:

1. **Copy the example configuration:**
   ```bash
   cp .env.test.example .env.test
   ```

2. **Edit `.env.test` with your values:**
   ```bash
   # Required for OAuth tests
   OAUTH_CLIENT_ID=your-client-id-here
   OAUTH_TENANT_ID=your-tenant-id-here
   OAUTH_REDIRECT_URI=https://your-domain.com/api/mcp/msgraph/oauth/callback
   
   # Server URLs
   LOCAL_URL=http://localhost:3001
   PUBLIC_REMOTE_URL=https://your-domain.com
   ```

3. **Load environment variables:**
   ```bash
   # Using dotenv
   npm install -g dotenv-cli
   dotenv -e .env.test -- node test-server.js local
   
   # Or export manually
   export OAUTH_CLIENT_ID=your-client-id
   export OAUTH_TENANT_ID=your-tenant-id
   # ... other variables
   ```

### Running Tests

```bash
# Install dependencies first
npm install

# Test local development server
npm test
# or
node test-server.js local

# Test remote Docker deployment (internal network)
npm run test:remote
# or  
node test-server.js remote

# Test public remote server
npm run test:public
# or
node test-server.js public

# Test custom URL
node test-server.js https://your-custom-url.com
```

### Test Categories

The test script covers:

1. **Health Endpoints** 🏥
   - `/health` - Server health check
   - `/` - Root endpoint with service info

2. **OAuth Endpoints** 🔐
   - `/.well-known/oauth-authorization-server` - OAuth discovery
   - `/register` - Dynamic client registration

3. **MCP Protocol** 🔧
   - `initialize` - MCP server initialization  
   - `tools/list` - Available tools listing
   - `ping` - Server ping/heartbeat

4. **MCP Tools** 🛠️
   - `throttling-stats` - API usage statistics
   - Error handling for invalid tools

5. **Error Handling** ⚠️
   - Invalid JSON requests
   - Unsupported methods
   - 404 endpoints

### Expected Results

#### Without Authentication
Most tests should pass, but tool calls requiring Microsoft Graph access will fail with authentication errors. This is expected behavior.

#### With Authentication  
All tests should pass when the server has valid OAuth tokens configured.

### Verbose Output

For detailed request/response logging:

```bash
VERBOSE=true node test-server.js local
```

## Manual Testing with LibreChat

### 1. Docker Setup
Ensure your MCP server is running in Docker alongside LibreChat:

```yaml
# docker-compose.yml excerpt
services:
  msgraph-mcp:
    build: .
    ports:
      - "3001:3001"
    environment:
      - TENANT_ID=b81c4237-e300-4913-abff-5bd60e9e2857
      - CLIENT_ID=151256ec-af20-432f-a064-09e0a10b6ab1
      - CLIENT_SECRET=your_client_secret
    networks:
      - librechat_network
```

### 2. LibreChat Configuration
Your `librechat.yaml` should include (replace with your actual values):

```yaml
mcpServers:
  msgraph:
    type: streamable-http
    url: http://msgraph-mcp:3001/mcp
    requiresOAuth: true
    startup: true
    oauth:
      client_id: your-client-id-here
      tenant_id: your-tenant-id-here
      redirect_uri: "https://your-domain.com/api/mcp/msgraph/oauth/callback"
      authorization_url: "https://login.microsoftonline.com/your-tenant-id/oauth2/v2.0/authorize"
      token_url: "https://login.microsoftonline.com/your-tenant-id/oauth2/v2.0/token"
      scope: "https://graph.microsoft.com/.default"
```

### 3. Testing OAuth Flow

1. **Start LibreChat** with the MCP server configured
2. **Navigate to a chat** where MCP tools are available
3. **Check connection** - LibreChat should show the MCP server as connected
4. **Test OAuth** - Try using a Microsoft Graph tool, which should prompt for authentication
5. **Verify tools** - After OAuth, tools like "Get my profile" should work

### 4. Testing Individual Tools

Try these MCP tools through LibreChat:

- **Get Profile**: `@msgraph get my user profile`
- **Search Files**: `@msgraph search for files containing "quarterly report"`
- **Get Calendar**: `@msgraph show my calendar events for today`
- **Get Schedule**: `@msgraph check my schedule from 9am to 5pm today`
- **Throttling Stats**: `@msgraph show API usage statistics`

## Troubleshooting

### Common Issues

1. **Connection Refused**
   - Check if server is running: `docker ps`
   - Verify port mapping: `3001:3001`
   - Check Docker network connectivity

2. **OAuth Errors**
   - Verify client credentials in environment variables
   - Check redirect URI matches exactly
   - Ensure tenant ID is correct

3. **Tool Call Failures**
   - Check server logs for detailed errors
   - Verify OAuth token is valid
   - Run `throttling-stats` tool to check API limits

4. **HTTP 500 Errors**
   - Check server logs: `docker logs msgraph-mcp`
   - Verify environment variables are set
   - Check Microsoft Graph API status

### Log Analysis

Enable verbose logging in your Docker environment:

```bash
# In docker-compose.yml
environment:
  - LOG_LEVEL=debug
  - VERBOSE=true
```

Then check logs:

```bash
docker logs -f msgraph-mcp
```

### Performance Testing

For load testing, you can run multiple test instances:

```bash
# Run 5 parallel test sessions
for i in {1..5}; do
  node test-server.js local &
done
wait
```

## CI/CD Integration

The test script returns appropriate exit codes for CI/CD:

```bash
# In your CI pipeline
npm install
npm run build
npm test

# Exit code 0 = all tests passed
# Exit code 1 = some tests failed
```
