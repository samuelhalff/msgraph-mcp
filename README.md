# Microsoft Graph MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/Node.js-18+-green.svg)](https://nodejs.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.6+-blue.svg)](https://www.typescriptlang.org/)
[![Docker](https://img.shields.io/badge/Docker-Ready-blue.svg)](https://www.docker.com/)

A powerful Microsoft Graph integration server built with the Model Context Protocol (MCP) for seamless integration with AI assistants like LibreChat. This server provides secure, authenticated access to Microsoft Graph APIs including users, groups, mail, calendar, OneDrive, Teams, and more.

## 🌟 Features

- **🔐 Secure OAuth 2.0 Authentication** with PKCE support
- **🤖 MCP Protocol Compliant** - Full Model Context Protocol implementation
- **🎯 LibreChat Ready** - Seamless integration with LibreChat and other MCP clients
- **📊 Comprehensive Microsoft Graph API Coverage**:
  - Users and Groups management
  - Mail and Calendar operations
  - OneDrive/SharePoint file operations with advanced search
  - Teams and Channels
  - Applications and Service Principals
  - Directory objects
- **⚡ Advanced File Search** - Content-based search across SharePoint and OneDrive
- **📅 Calendar Integration** - Schedule management and availability checking
- **🛡️ Intelligent Throttling** - Built-in rate limiting with exponential backoff
- **🐳 Docker Ready** - Optimized container with health checks
- **📝 Structured Logging** - Winston-based logging with file and console output
- **🔄 Token Refresh** - Automatic token refresh handling
- **🛡️ Multiple Authentication Methods**:
  - Client Credentials
  - Interactive Browser
  - Certificate-based authentication
  - Client-provided tokens

## 🚀 Quick Start

### Prerequisites

- Node.js 18+ and npm
- Microsoft Azure App Registration
- Docker (optional, for containerized deployment)

### 1. Clone and Install

```bash
git clone https://github.com/samuelhalff/msgraph-mcp.git
cd msgraph-mcp
npm install
```

### 2. Configure Microsoft Azure

1. Go to [Azure Portal](https://portal.azure.com) → App registrations
2. Create a new app registration or use existing one
3. Note down your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET`
4. Configure redirect URI: `http://localhost:3001/auth/callback`
5. Grant necessary Microsoft Graph permissions (e.g., `https://graph.microsoft.com/.default`)

### 3. Environment Setup

```bash
cp .env.example .env
```

Edit `.env` with your Azure credentials:

```bash
# Microsoft Graph App Registration Settings
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here
REDIRECT_URI=http://localhost:3001/auth/callback

# Microsoft Graph Configuration
OAUTH_SCOPES=https://graph.microsoft.com/.default
USE_GRAPH_BETA=false
USE_INTERACTIVE=false
USE_CLIENT_TOKEN=true
USE_CERTIFICATE=false

# Server Configuration
PORT=3001
```

### 4. Build and Run

```bash
# Development mode
npm run dev

# Production build
npm run build
npm start

# Or using Docker
docker-compose up -d
```

## 📖 API Endpoints

### OAuth Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/.well-known/oauth-authorization-server` | GET | OAuth discovery document |
| `/register` | POST | Client registration |
| `/authorize` | GET | OAuth authorization (redirects to Microsoft) |
| `/token` | POST | Token exchange |
| `/userinfo` | GET | User information (requires auth) |
| `/logout` | POST | Logout |

### MCP Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/mcp` | POST | MCP protocol endpoint |
| `/health` | GET | Health check |

### MCP Tools Available

- **microsoft-graph-api**: Generic Microsoft Graph API access
- **microsoft-graph-profile**: Get current user profile
- **list-users**: List organization users
- **list-groups**: List organization groups
- **search-users**: Search for users by name
- **send-mail**: Send emails through Microsoft Graph
- **search-files**: Advanced file search across SharePoint and OneDrive
- **get-schedule**: Calendar and scheduling operations
- **throttling-stats**: Monitor API performance and rate limiting

## 🎯 LibreChat Integration

This MCP server is designed for seamless integration with [LibreChat](https://www.librechat.ai/), providing powerful Microsoft 365 capabilities to AI assistants.

### Quick LibreChat Setup

1. **Deploy this MCP server** (see deployment instructions above)
2. **Add to your `librechat.yaml`**:

```yaml
mcpServers:
  msgraph:
    type: "streamable-http"
    url: "https://your-server.com/mcp"
    initTimeout: 30000
    timeout: 60000
    headers:
      Authorization: "Bearer {{MS_GRAPH_TOKEN}}"
    customUserVars:
      MS_GRAPH_TOKEN:
        title: "Microsoft Graph Access Token"
        description: "Get your token at https://your-server.com/authorize"
    serverInstructions: |
      Microsoft Graph integration for Microsoft 365:
      - Use search-files for SharePoint/OneDrive document search
      - Use get-schedule for calendar management
      - Use send-mail for email operations
    startup: false
```

3. **Restart LibreChat** and authenticate with Microsoft

### Features in LibreChat

- 📁 **Smart File Search**: Ask "Find PowerPoint files about quarterly reports"
- 📧 **Email Management**: "Send a project update to the team"  
- 📅 **Calendar Integration**: "Check my availability next Tuesday"
- 👥 **User Directory**: "Find contact info for Sarah in marketing"
- 📊 **Performance Monitoring**: Built-in throttling protection

**📋 For complete LibreChat integration guide, see [LIBRECHAT.md](./LIBRECHAT.md)**

## 🔧 Configuration

### Environment Variables

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `TENANT_ID` | Azure tenant ID | Yes | - |
| `CLIENT_ID` | Azure app client ID | Yes | - |
| `CLIENT_SECRET` | Azure app client secret | Yes* | - |
| `REDIRECT_URI` | OAuth redirect URI | Yes | - |
| `OAUTH_SCOPES` | Microsoft Graph scopes | No | `https://graph.microsoft.com/.default` |
| `USE_GRAPH_BETA` | Use Graph beta endpoint | No | `false` |
| `USE_CLIENT_TOKEN` | Use client credentials auth | No | `true` |
| `USE_INTERACTIVE` | Use interactive browser auth | No | `false` |
| `USE_CERTIFICATE` | Use certificate auth | No | `false` |
| `CERTIFICATE_PATH` | Path to certificate file | No* | - |
| `CERTIFICATE_PASSWORD` | Certificate password | No* | - |
| `PORT` | Server port | No | `3001` |
| `LOG_LEVEL` | Logging level | No | `info` |

*Required only when using the respective authentication method

### Authentication Methods

#### 1. Client Credentials (Recommended)
```bash
USE_CLIENT_TOKEN=true
USE_INTERACTIVE=false
USE_CERTIFICATE=false
```

#### 2. Interactive Browser
```bash
USE_CLIENT_TOKEN=false
USE_INTERACTIVE=true
USE_CERTIFICATE=false
```

#### 3. Certificate-Based
```bash
USE_CLIENT_TOKEN=false
USE_INTERACTIVE=false
USE_CERTIFICATE=true
CERTIFICATE_PATH=/path/to/cert.pem
CERTIFICATE_PASSWORD=your-password
```

## 🐳 Docker Deployment

### Using Docker Compose

```bash
# Build and run
docker-compose up -d

# View logs
docker-compose logs -f msgraph-mcp

# Stop services
docker-compose down
```

### Manual Docker Build

```bash
# Build image
docker build -t msgraph-mcp .

# Run container
docker run -d \
  --name msgraph-mcp \
  -p 3001:3001 \
  --env-file .env \
  msgraph-mcp
```

## 🔗 LibreChat Integration

### MCP Configuration

Add to your LibreChat MCP configuration:

```yaml
msgraph:
  type: streamable-http
  url: http://localhost:3001/mcp
  oauth:
    discovery_url: http://localhost:3001/.well-known/oauth-authorization-server
```

### Docker Network Setup

When running with LibreChat in Docker:

```yaml
msgraph:
  type: streamable-http
  url: http://msgraph-mcp:3001/mcp
  oauth:
    discovery_url: http://msgraph-mcp:3001/.well-known/oauth-authorization-server
```

## 📊 Logging

The server provides comprehensive logging:

- **Console Output**: Real-time logs with colors (development)
- **File Logs**: Persistent logs in `logs/` directory
  - `logs/combined.log`: All logs
  - `logs/error.log`: Error and warning logs only

### Log Levels

Set `LOG_LEVEL` environment variable:
- `error`: Errors only
- `warn`: Warnings and errors
- `info`: General information (default)
- `debug`: Detailed debugging

## 🛠️ Development

### Project Structure

```
msgraph-mcp/
├── api/                    # Source code
│   ├── lib/
│   │   ├── logger.ts      # Winston logger configuration
│   │   └── msgraph-auth.ts # Authentication helpers
│   ├── MSGraphMCP.ts      # MCP server implementation
│   ├── MSGraphService.ts  # Microsoft Graph client
│   └── index.ts           # Main server entry point
├── dist/                   # Compiled JavaScript
├── logs/                   # Log files
├── public/                 # Static assets
├── server.js              # Production server
├── Dockerfile             # Docker configuration
├── docker-compose.yml     # Docker Compose setup
└── package.json          # Dependencies and scripts
```

### Development Commands

```bash
# Start development server with hot reload
npm run dev

# Build for production
npm run build

# Start production server
npm start

# Lint code
npm run lint
```

### Adding New Microsoft Graph Features

1. Add methods to `MSGraphService.ts`
2. Update MCP tools in `MSGraphMCP.ts`
3. Add corresponding TypeScript types in `types.d.ts`
4. Test with LibreChat integration

## 🔍 Troubleshooting

### Common Issues

#### 401 Unauthorized
- Check Azure app permissions
- Verify `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`
- Ensure correct redirect URI is configured

#### Connection Refused
- Verify server is running on correct port
- Check Docker network configuration
- Ensure MCP URL is accessible from LibreChat

#### OAuth Flow Issues
- Check OAuth discovery endpoint: `/.well-known/oauth-authorization-server`
- Verify PKCE configuration
- Check Azure app registration settings

#### Permission Errors
- Ensure `logs/` directory exists and is writable
- Check Docker user permissions
- Verify file system permissions

### Debug Mode

Enable detailed logging:

```bash
LOG_LEVEL=debug npm start
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit changes: `git commit -am 'Add your feature'`
4. Push to branch: `git push origin feature/your-feature`
5. Submit a pull request

### Development Guidelines

- Use TypeScript for all new code
- Follow existing code style and patterns
- Add comprehensive logging for new features
- Update documentation for API changes
- Test with both local and Docker environments

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Microsoft Graph](https://docs.microsoft.com/en-us/graph/) - RESTful web API for Microsoft services
- [Model Context Protocol](https://modelcontextprotocol.io/) - Open standard for tool use
- [LibreChat](https://github.com/danny-avila/LibreChat) - Open-source chat interface
- [Hono](https://hono.dev/) - Fast web framework for Cloudflare Workers and more

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/samuelhalff/msgraph-mcp/issues)
- **Discussions**: [GitHub Discussions](https://github.com/samuelhalff/msgraph-mcp/discussions)
- **Documentation**: [Microsoft Graph Docs](https://docs.microsoft.com/en-us/graph/)

---

**Made with ❤️ for seamless Microsoft Graph integration with AI assistants**