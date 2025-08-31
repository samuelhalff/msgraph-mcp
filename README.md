# Microsoft Graph MCP Server

A Model Context Protocol (MCP) server that provides Microsoft Graph API integration for LibreChat and other MCP-compatible applications.

## Features

- 🔐 **OAuth 2.0 Authentication** - Secure authentication with Microsoft Graph
- 📧 **Email Operations** - Send, read, and manage emails
- 👥 **User Management** - Access user profiles and directory information
- 📅 **Calendar Integration** - Manage events and schedules
- 📁 **OneDrive Integration** - File storage and sharing
- 👥 **Teams Integration** - Microsoft Teams functionality
- 🏢 **Administrative APIs** - Organization and tenant management

## Quick Start

### 1. Clone and Install

```bash
git clone <repository-url>
cd msgraph-mcp
npm install
```

### 2. Microsoft Graph App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **App registrations** → **New registration**
3. Configure:
   - Name: `LibreChat MS Graph MCP`
   - Supported account types: `Accounts in any organizational directory`
   - Redirect URI: `http://localhost:3001/auth/callback`

4. Note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

5. Create a **Client Secret**:
   - Go to **Certificates & secrets**
   - **New client secret**
   - Note down the **Value** (not the Secret ID)

6. Configure **API Permissions**:
   - **Microsoft Graph** → **Delegated permissions**
   - Add: `https://graph.microsoft.com/.default`

### 3. Environment Configuration

```bash
cp .env.example .env
```

Edit `.env` with your values:
```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
REDIRECT_URI=http://localhost:3001/auth/callback
OAUTH_SCOPES=https://graph.microsoft.com/.default
```

### 4. Start the Server

```bash
# Development mode
npm run dev

# Production mode
npm run build
npm start
```

The server will start on `http://localhost:3001`

### 5. Configure LibreChat

Add to your `librechat.yaml`:

```yaml
mcpServers:
  msgraph:
    type: streamable-http
    url: http://localhost:3001/mcp
    requiresOAuth: true
    env:
      TENANT_ID: "your-tenant-id"
      CLIENT_ID: "your-client-id"
      CLIENT_SECRET: "your-client-secret"
      REDIRECT_URI: "http://localhost:3001/auth/callback"
      SCOPE: "https://graph.microsoft.com/.default"
```

## Docker Deployment

### Option 1: Standalone Docker

```bash
# Build and run
docker build -t msgraph-mcp .
docker run -p 3001:3001 --env-file .env msgraph-mcp
```

### Option 2: Docker Compose (with LibreChat)

```bash
# Copy docker-compose.yml to your LibreChat directory
cp docker-compose.yml /path/to/librechat/

# Start both services
docker-compose up -d
```

## API Endpoints

- `GET /.well-known/oauth-authorization-server` - OAuth discovery
- `POST /register` - Client registration
- `GET /authorize` - Authorization redirect
- `POST /token` - Token exchange
- `GET /userinfo` - User information
- `POST /logout` - Logout
- `POST /mcp` - MCP protocol endpoint
- `GET /health` - Health check

## Available Tools

The MCP server provides these tools to LibreChat:

### Core Tools
- `microsoft-graph-api` - Versatile Microsoft Graph API access
- `get-auth-status` - Check authentication status

### Microsoft Graph Tools
- `getUsers` - List users
- `getGroups` - List groups
- `getApplications` - List applications
- `getCurrentUserProfile` - Get current user profile
- `sendMail` - Send emails
- `getMessages` - Read emails
- `createEvent` - Create calendar events
- `getEvents` - List calendar events
- `getDriveItems` - Access OneDrive files
- `getTeams` - List Microsoft Teams
- `getChannels` - List team channels

## Development

```bash
# Install dependencies
npm install

# Development server with hot reload
npm run dev

# Build for production
npm run build

# Run production server
npm start

# Run linting
npm run lint
```

## Configuration Options

| Environment Variable | Description | Default |
|---------------------|-------------|---------|
| `TENANT_ID` | Azure tenant ID | Required |
| `CLIENT_ID` | Azure app client ID | Required |
| `CLIENT_SECRET` | Azure app client secret | Required |
| `REDIRECT_URI` | OAuth redirect URI | Required |
| `OAUTH_SCOPES` | Microsoft Graph scopes | `https://graph.microsoft.com/.default` |
| `USE_GRAPH_BETA` | Use Graph beta endpoint | `false` |
| `USE_CLIENT_TOKEN` | Use client-provided tokens | `true` |
| `PORT` | Server port | `3001` |

## Troubleshooting

### Common Issues

1. **"Invalid client" error**
   - Check your CLIENT_ID and CLIENT_SECRET
   - Ensure the app registration has correct permissions

2. **"Invalid scope" error**
   - Verify OAUTH_SCOPES includes required permissions
   - Check app registration API permissions

3. **Connection refused**
   - Ensure the server is running on the correct port
   - Check firewall settings

### Logs

```bash
# View server logs
docker logs msgraph-mcp

# View LibreChat logs
docker logs librechat
```

## Security Notes

- Store secrets securely (use Docker secrets or environment variables)
- Use HTTPS in production
- Regularly rotate client secrets
- Limit API permissions to only what's needed

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## Support

- 📖 [Documentation](https://www.librechat.ai/docs)
- 🐛 [Issues](https://github.com/samuelhalff/msgraph-mcp/issues)
- 💬 [Discussions](https://github.com/samuelhalff/msgraph-mcp/discussions)

### Not Production-Ready

Before deploying to production, you should:

1. **Add Authentication Validation**
   - Implement proper client validation in the `/register` endpoint
   - Add rate limiting to prevent abuse
   - Validate redirect URIs against a whitelist

2. **Enhance Security**
   - Implement CSRF protection
   - Add request signing/verification
   - Use secure session management
   - Implement proper error handling that doesn't leak sensitive information

3. **Add Monitoring & Logging**
   - Implement comprehensive logging
   - Add error tracking (e.g., Sentry)
   - Monitor API usage and rate limits
   - Track OAuth flow completion rates

4. **Optimize for Scale**
   - Implement caching strategies
   - Add database for persistent storage (instead of KV for production data)
   - Implement proper connection pooling
   - Add request queuing for rate-limited operations

5. **Handle Edge Cases**
   - Implement proper retry logic with exponential backoff
   - Handle Spotify API maintenance windows
   - Add graceful degradation for non-critical features
   - Implement proper timeout handling

### What's Already Production-Ready

Thanks to the original Stytch implementation, this project already includes:

✅ **Cloudflare Workers Deployment**
- Full Workers configuration with `wrangler.jsonc`
- Durable Objects setup for MCP server instances
- Workers KV integration for data storage
- Proper TypeScript configuration for Workers environment

✅ **Infrastructure Foundation**
- SSE (Server-Sent Events) implementation for real-time MCP communication
- Proper CORS handling
- Environment variable management
- Build and deployment scripts

This means you already have a solid infrastructure foundation - you just need to add the security and monitoring layers mentioned above for production use.

### Production Deployment Checklist

If you plan to use this in production:
- [ ] Review and implement all security considerations
- [ ] Add comprehensive error handling
- [ ] Implement proper logging and monitoring
- [ ] Add automated tests
- [ ] Set up CI/CD pipeline
- [ ] Configure production environment variables securely
- [ ] Implement backup and recovery procedures
- [ ] Add API versioning strategy
- [ ] Create documentation for your specific implementation
- [ ] Implement user consent and data privacy compliance

## MCP Inspector Compatibility

Following the pattern from [PayPal's MCP server](https://developer.paypal.com/tools/mcp-server/), this implementation provides:

1. **OAuth Discovery** at `/.well-known/oauth-authorization-server`
2. **Dynamic Client Registration** at `/register` endpoint
3. **Authorization Proxy** that redirects to Spotify's OAuth system

This allows MCP Inspector and other MCP clients to automatically discover and register with our server, even though Spotify itself doesn't support Dynamic Client Registration.

## Setup

### 1. Create a Spotify App

1. Go to [Spotify Developer Dashboard](https://developer.spotify.com/dashboard)
2. Create a new app
3. Note your Client ID and Client Secret
4. Add redirect URIs for your application (e.g., `http://localhost:3000/callback`)

### 2. Configure Environment Variables

Create a `.dev.vars` file based on `.dev.vars.template`:

```
SPOTIFY_CLIENT_ID=your_spotify_client_id
SPOTIFY_CLIENT_SECRET=your_spotify_client_secret
```

### 3. Install Dependencies

```bash
npm install
```

### 4. Run Locally

```bash
npm run dev
```

The MCP server will be available at `http://localhost:3000/sse`

### Environment Variables for Vite

The `vite.config.ts` file supports additional environment variables for development:

- **`VITE_PORT`**: Customize the development server port (default: 3000)
  ```bash
  VITE_PORT=8080 npm run dev
  ```

- **`VITE_ALLOWED_HOSTS`**: Specify allowed hosts for the dev server (useful when using tunneling services)
  ```bash
  VITE_ALLOWED_HOSTS="localhost,your-ngrok-domain.ngrok-free.app" npm run dev
  ```

## HTTPS Requirements

### The Problem

The Spotify Web API requires HTTPS for OAuth callbacks, and modern browsers enforce secure connections between MCP clients and servers to prevent mixed content issues.

### Solutions by Deployment Type

#### Production Deployment (Recommended)

**✅ Cloudflare Workers Solution**
- **MCP Server**: Deploy to Cloudflare Workers - automatic HTTPS included
  - Your deployed server: `https://your-worker.workers.dev`
  - Permanent, stable HTTPS URLs
  - No additional SSL configuration required
  - Global edge network for fast response times

- **MCP Client**: Most production MCP clients already run over HTTPS
- **Spotify App**: Update redirect URIs to use your Workers domain (one-time setup)

```bash
# Deploy to get permanent HTTPS URL
npm run deploy

# Update Spotify app redirect URI to:
# https://your-worker.workers.dev/callback
```

#### Local Development

**🔧 ngrok Solution for Local Testing**

When developing locally, use [ngrok](https://ngrok.com/) to create HTTPS tunnels:

##### 1. Install ngrok
```bash
# Download from https://ngrok.com/download or use package manager
brew install ngrok  # macOS
# or
snap install ngrok  # Linux
```

##### 2. Start Your Local MCP Server
```bash
# Start with allowed hosts configuration
VITE_ALLOWED_HOSTS="localhost,*.ngrok-free.app" npm run dev
```

##### 3. Create HTTPS Tunnels

For local development, you may need tunnels for both server and client:

```bash
# Terminal 1: Tunnel for MCP Server (port 3000)
ngrok http 3000

# Terminal 2: Tunnel for LibreChat if running locally (port 3080)
ngrok http 3080
```

##### 4. Configure for Development Session

After starting ngrok, you'll get temporary HTTPS URLs:
- MCP Server: `https://abc123.ngrok-free.app`
- LibreChat: `https://def456.ngrok-free.app`

Update your Spotify app for this development session:
1. Go to [Spotify Developer Dashboard](https://developer.spotify.com/dashboard)
2. Add the ngrok redirect URI: `https://abc123.ngrok-free.app/callback`
3. Configure MCP client to use: `https://abc123.ngrok-free.app/sse`

### Comparison: Production vs Development

| Aspect | Cloudflare Workers (Production) | ngrok (Development) |
|--------|--------------------------------|---------------------|
| **HTTPS** | ✅ Automatic, permanent | ✅ Temporary tunnels |
| **URLs** | ✅ Stable, persistent | ❌ Change on restart |
| **Setup** | ✅ One-time deployment | ❌ Per-session setup |
| **Cost** | ✅ Free tier available | ✅ Free tier available |
| **Performance** | ✅ Global edge network | ❌ Tunneling overhead |
| **Use Case** | Production, permanent testing | Local development |

### Recommended Workflow

1. **Start with Production**: Deploy to Cloudflare Workers first for stable testing
2. **Use Local When Needed**: Use ngrok only when you need to test local code changes
3. **Hybrid Approach**: Keep production deployment for stable testing, use local for active development

### ngrok Limitations (Development Only)

- **Temporary Domains**: URLs change each restart (requires updating Spotify app)
- **Session Limits**: 2-hour sessions on free plan
- **Request Limits**: 40 requests per minute on free plan
- **Multiple Tunnels**: Need separate ngrok instances for server and client

### Production Considerations

**✅ Cloudflare Workers Benefits:**
- Automatic HTTPS with valid certificates
- Global CDN for low latency
- Built-in DDoS protection
- 99.9% uptime SLA
- Easy custom domain support

**For Enterprise Use:**
- Custom domains: `mcp.yourcompany.com`
- Enhanced security with Cloudflare WAF
- Advanced analytics and monitoring
- Team collaboration features

## OAuth Flow

### How It Works with MCP Inspector

1. **Discovery**: MCP Inspector fetches `/.well-known/oauth-authorization-server`
2. **Registration**: Inspector registers itself at `/register` endpoint
3. **Authorization**: User is redirected to `/authorize`, which redirects to Spotify
4. **User Consent**: User authorizes on Spotify's page
5. **Token Exchange**: Inspector exchanges code at `/token`
6. **Access MCP**: Inspector uses tokens to access MCP endpoints via SSE

## Available MCP Tools

The Spotify MCP server exposes the following tools:

### Search
- `searchTracks` - Search for tracks
- `searchArtists` - Search for artists
- `searchAlbums` - Search for albums
- `searchPlaylists` - Search for playlists

### User Profile
- `getCurrentUserProfile` - Get current user's profile

### Playback Control
- `getCurrentPlayback` - Get current playback state
- `pausePlayback` - Pause playback
- `resumePlayback` - Resume playback
- `skipToNext` - Skip to next track
- `skipToPrevious` - Skip to previous track

### Playlists
- `getUserPlaylists` - Get user's playlists
- `getPlaylistTracks` - Get tracks from a playlist
- `createPlaylist` - Create a new playlist
- `addTracksToPlaylist` - Add tracks to a playlist

### User Data
- `getRecentlyPlayed` - Get recently played tracks
- `getTopTracks` - Get user's top tracks
- `getTopArtists` - Get user's top artists

## Example Usage

### Using MCP Inspector

1. Open MCP Inspector:
   ```bash
   npx @modelcontextprotocol/inspector@latest
   ```
2. Set Transport Type to `SSE`
3. Enter URL: `http://localhost:3000/sse`
4. Click Connect
5. Follow the OAuth flow when redirected

### Manual Token Usage

If you prefer to bypass the OAuth flow for testing:

1. Get tokens using the manual flow
2. In MCP Inspector, add headers:
   - `Authorization: Bearer YOUR_ACCESS_TOKEN`
   - `X-Spotify-Refresh-Token: YOUR_REFRESH_TOKEN`

## Deployment

### Quick Deploy to Cloudflare Workers

**🚀 Automatic HTTPS Solution**

Deploying to Cloudflare Workers automatically solves the HTTPS requirement:

```bash
# 1. Set your Spotify credentials
wrangler secret put SPOTIFY_CLIENT_ID
wrangler secret put SPOTIFY_CLIENT_SECRET

# 2. Deploy (gets automatic HTTPS)
npm run deploy

# 3. Update Spotify app redirect URI to your new Workers URL
# Example: https://your-worker.workers.dev/callback
```

Your MCP server will be available at: `https://your-worker.workers.dev/sse`

**⚠️ Remember: This is a development example. See the "Important: Development Example" section above for production considerations.**

### Production Deployment

For actual production use:
1. Fork this repository
2. Implement the security enhancements listed above
3. Add proper monitoring and error handling
4. Consider using Cloudflare's paid features:
   - Durable Objects for better state management
   - Workers KV for caching
   - Rate limiting rules
   - Web Application Firewall (WAF)

## Security Considerations

1. **Client Credentials**: Never expose your Spotify Client Secret in client-side code
2. **Token Storage**: Store tokens securely and use HTTPS for all communications
3. **Scopes**: Only request the minimum scopes necessary for your application
4. **PKCE**: The implementation supports PKCE for added security
5. **Client Registration**: In production, consider adding validation to the registration endpoint

## Spotify API Gotchas

### 1. HTTPS Requirement
- Spotify Web API requires HTTPS for OAuth callbacks
- **Solution**: Deploy to Cloudflare Workers for automatic HTTPS
- **Local Development**: Use ngrok for HTTPS tunneling
- Mixed content errors occur when client and server use different protocols

### 2. Active Device Requirement
- Playback control endpoints (`pausePlayback`, `resumePlayback`, etc.) require an active Spotify device
- Users must have Spotify open on at least one device (desktop app, mobile, or web player)
- API returns 404 "No active device found" if no device is available

### 3. Premium Account Limitations
- Many playback control features require a Spotify Premium account
- Free accounts can read data but cannot control playback
- The API returns 403 "Premium required" errors for free accounts

### 4. Rate Limiting
- Spotify implements rate limiting on all endpoints
- Limits vary by endpoint but typically allow 180 requests per minute
- Implement exponential backoff for 429 (Too Many Requests) responses

### 5. Token Expiration
- Access tokens expire after 1 hour
- This implementation handles automatic refresh, but ensure refresh tokens are stored
- Refresh tokens can become invalid if unused for extended periods

### 6. Scope Requirements
- Different endpoints require different OAuth scopes
- Missing scopes result in 403 Forbidden errors
- Common required scopes:
  - `user-read-private` - User profile access
  - `user-read-playback-state` - Current playback info
  - `user-modify-playback-state` - Playback control
  - `playlist-read-private` - Access user playlists
  - `playlist-modify-public` - Create/modify playlists

### 7. Market Restrictions
- Some content is restricted by geographic market
- API may return different results based on user's country
- Use the `market` parameter when searching to get region-appropriate results

## Acknowledgments

This project is based on the excellent [Stytch MCP Consumer TODO List](https://github.com/stytchauth/mcp-stytch-consumer-todo-list) example, which demonstrated how to implement OAuth discovery for MCP servers. We've adapted their pattern to work with Spotify's OAuth system.

## Resources

- [Spotify Web API Documentation](https://developer.spotify.com/documentation/web-api/)
- [OAuth 2.0 Authorization Code Flow](https://developer.spotify.com/documentation/web-api/tutorials/code-flow)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Cloudflare Workers Documentation](https://developers.cloudflare.com/workers/)
- [Original Stytch Example](https://github.com/stytchauth/mcp-stytch-consumer-todo-list)

