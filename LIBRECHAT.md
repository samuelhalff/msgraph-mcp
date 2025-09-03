# LibreChat Integration Guide

This guide explains how to integrate the Microsoft Graph MCP Server with LibreChat for seamless Microsoft 365 integration.

## Overview

This Microsoft Graph MCP server provides LibreChat users with comprehensive Microsoft 365 functionality including:
- User and group management
- Email sending capabilities
- File search across SharePoint and OneDrive
- Calendar and scheduling operations
- Full Microsoft Graph API access

## LibreChat Configuration

### Basic Configuration

Add the following to your `librechat.yaml` file:

```yaml
mcpServers:
  msgraph:
    type: "streamable-http"
    url: "https://your-server.com/mcp"
    initTimeout: 30000
    timeout: 60000
    headers:
      Authorization: "Bearer {{MS_GRAPH_TOKEN}}"
      X-User-ID: "{{LIBRECHAT_USER_ID}}"
      X-User-Email: "{{LIBRECHAT_USER_EMAIL}}"
    customUserVars:
      MS_GRAPH_TOKEN:
        title: "Microsoft Graph Access Token"
        description: "Your Microsoft Graph API access token. <a href='https://your-server.com/authorize' target='_blank'>Get token here</a>"
    serverInstructions: |
      Microsoft Graph integration for Microsoft 365:
      - Use search-files for finding documents in SharePoint/OneDrive
      - Use get-schedule for calendar and meeting management
      - Use send-mail for email operations
      - Always check throttling-stats if requests seem slow
    startup: false  # Requires user authentication first
```

### OAuth-Enabled Configuration (Recommended)

For production environments with OAuth support:

```yaml
mcpServers:
  msgraph-oauth:
    type: "streamable-http"
    url: "https://your-msgraph-server.com/mcp"
    initTimeout: 150000  # Higher timeout for OAuth flow
    timeout: 60000
    headers:
      X-User-ID: "{{LIBRECHAT_USER_ID}}"
      X-User-Email: "{{LIBRECHAT_USER_EMAIL}}"
      X-User-Role: "{{LIBRECHAT_USER_ROLE}}"
    serverInstructions: |
      Microsoft Graph OAuth integration:
      - Automatically handles authentication with Microsoft 365
      - Full access to user's Microsoft Graph data
      - Supports file search, calendar, email, and user management
      - Includes automatic throttling and retry logic
    chatMenu: true  # Available in chat dropdown
```

## Server Deployment

### Environment Variables

Configure these environment variables for your MCP server:

```bash
# Microsoft Graph App Registration (Required)
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret

# Server Configuration
PORT=3001
PUBLIC_BASE_URL=https://your-server.com
REDIRECT_URI=https://your-server.com/callback

# Optional: Advanced Configuration
USE_CERTIFICATE=false
CERTIFICATE_PATH=/path/to/cert.pem
CERTIFICATE_PASSWORD=cert-password
```

### Docker Deployment

```dockerfile
FROM node:18-alpine

WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production

COPY . .
RUN npm run build

EXPOSE 3001

ENV NODE_ENV=production
CMD ["npm", "start"]
```

### Docker Compose Example

```yaml
version: '3.8'
services:
  msgraph-mcp:
    build: .
    ports:
      - "3001:3001"
    environment:
      - TENANT_ID=${TENANT_ID}
      - CLIENT_ID=${CLIENT_ID}
      - CLIENT_SECRET=${CLIENT_SECRET}
      - PUBLIC_BASE_URL=https://your-domain.com
      - REDIRECT_URI=https://your-domain.com/callback
      - NODE_ENV=production
    restart: unless-stopped
    
  # Reverse proxy (nginx/traefik) recommended for HTTPS
  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf
      - ./ssl:/etc/nginx/ssl
    depends_on:
      - msgraph-mcp
```

## Authentication Setup

### Microsoft App Registration

1. Go to [Azure Portal > App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click "New registration"
3. Configure:
   - **Name**: "LibreChat Microsoft Graph MCP"
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: `https://your-server.com/callback`

4. After creation, note the **Application (client) ID** and **Directory (tenant) ID**
5. Go to **Certificates & secrets** > **New client secret** > Note the secret value
6. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**

### Required Permissions

Add these Microsoft Graph permissions:

**Delegated Permissions:**
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read all users' basic profiles  
- `Group.Read.All` - Read all groups
- `Mail.Send` - Send emails as user
- `Files.Read.All` - Read all files user can access
- `Files.ReadWrite.All` - Read and write all files user can access
- `Calendars.Read` - Read user's calendars
- `Calendars.ReadBasic` - Read basic calendar info for all users

**Application Permissions (for service scenarios):**
- `User.Read.All` - Read all users' profiles
- `Group.Read.All` - Read all groups
- `Files.Read.All` - Read all files in organization
- `Calendars.Read` - Read calendars in all mailboxes

### Grant Admin Consent

1. In App registration > **API permissions**
2. Click **Grant admin consent for [Your Organization]**
3. Confirm the action

## Available Tools

Once configured, LibreChat users will have access to these Microsoft Graph tools:

### Core Tools
- **microsoft-graph-api**: Direct Graph API access
- **microsoft-graph-profile**: Get current user profile
- **list-users**: List organization users
- **list-groups**: List organization groups
- **search-users**: Search for users by name

### Communication Tools
- **send-mail**: Send emails through Microsoft Graph
- **get-schedule**: Get calendar/scheduling information

### File Management Tools
- **search-files**: Search files in SharePoint and OneDrive
  - Supports content search, file type filtering
  - Returns direct download/view links
  - Includes metadata and permissions

### Monitoring Tools
- **throttling-stats**: Monitor API performance and throttling

## Usage Examples

### In LibreChat Chat

1. Select a compatible model (GPT-4, Claude, etc.)
2. Click the tools dropdown below the message input
3. Select "Microsoft Graph MCP" server
4. Ask questions like:
   - "Search for PowerPoint files about quarterly reports"
   - "Schedule a meeting with the marketing team next week"
   - "Send an email to john@company.com about the project update"
   - "Find all Excel files modified in the last week"

### With LibreChat Agents

1. Go to Agent Builder
2. Click "Add Tools"
3. Select "Microsoft Graph MCP" server
4. Choose specific tools or enable all
5. Save your agent with Microsoft 365 capabilities

## Security Considerations

### Token Management
- Tokens are securely stored and automatically refreshed
- Each user maintains their own authentication session
- No long-lived credentials stored in configuration

### Permissions
- Users only access data they have permission to see
- Respects Microsoft 365 security boundaries
- Supports both delegated and application permissions

### Rate Limiting
- Built-in throttling protection with exponential backoff
- Respects Microsoft Graph service limits
- Automatic retry logic for transient failures

## Troubleshooting

### Connection Issues

1. **Server not connecting**: Check if server is accessible at the configured URL
2. **Authentication failures**: Verify Azure app registration and permissions
3. **Rate limiting**: Monitor throttling-stats tool for API usage

### Common Error Messages

**"MCP request failed"**
- Check server logs for detailed error information
- Verify environment variables are correctly set
- Ensure Microsoft Graph permissions are granted

**"OAuth authentication required"**
- User needs to authenticate with Microsoft
- Click the authentication button in LibreChat's MCP panel
- Complete OAuth flow in browser

**"Throttling detected"**
- Microsoft Graph API limits exceeded
- Server will automatically retry with backoff
- Check throttling-stats for current status

### Debug Mode

Enable detailed logging by setting:
```bash
NODE_ENV=development
DEBUG=msgraph-mcp:*
```

## Advanced Configuration

### Custom Scopes

Modify the OAuth scope in your server configuration:

```typescript
// In msgraph-auth.ts
const defaultScope = [
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Files.Read.All',
  'https://graph.microsoft.com/Mail.Send',
  'https://graph.microsoft.com/Calendars.Read'
].join(' ');
```

### Multi-Tenant Support

For multi-tenant scenarios, configure:

```bash
TENANT_ID=common  # Supports any organizational tenant
# or
TENANT_ID=organizations  # Any organizational account
```

### Certificate Authentication

For enhanced security in production:

```bash
USE_CERTIFICATE=true
CERTIFICATE_PATH=/app/certificates/msgraph.pem
CERTIFICATE_PASSWORD=your-cert-password
```

## Support

For issues specific to LibreChat integration:
- [LibreChat Documentation](https://www.librechat.ai/docs/features/mcp)
- [LibreChat Discord](https://discord.librechat.ai/)

For Microsoft Graph MCP server issues:
- Check server logs for detailed error information
- Review Microsoft Graph API documentation
- Verify Azure app registration configuration
