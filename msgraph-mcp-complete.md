# Microsoft Graph MCP Server - Complete Implementation

This is a complete, functional TypeScript MCP server for Microsoft Graph integration with LibreChat. LibreChat handles OAuth authentication and passes tokens to the MCP server.

## Setup Instructions

1. Save this file as `msgraph-mcp-complete.md`
2. Run the extraction script: `./creation-script.sh msgraph-mcp-complete.md`
3. Configure your Azure app registration
4. Copy `.env.example` to `.env` and fill in your values
5. Build and run: `npm install && npm run build && npm start`

---

## filename: package.json
```json
{
  "name": "msgraph-mcp-server",
  "version": "1.0.0",
  "description": "Microsoft Graph MCP Server for LibreChat with OAuth2 support",
  "main": "build/index.js",
  "scripts": {
    "build": "tsc",
    "start": "node build/index.js",
    "dev": "tsc && node build/index.js"
  },
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.0.0",
    "@azure/msal-node": "^2.6.6",
    "express": "^4.18.2",
    "dotenv": "^16.3.1",
    "axios": "^1.6.0",
    "uuid": "^9.0.0",
    "zod": "^3.22.4"
  },
  "devDependencies": {
    "@types/express": "^4.17.21",
    "@types/node": "^20.0.0",
    "@types/uuid": "^9.0.0",
    "typescript": "^5.0.0"
  }
}
```

## filename: tsconfig.json
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "commonjs",
    "outDir": "./build",
    "rootDir": "./src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "resolveJsonModule": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "build"]
}
```

## filename: .env.example
```env
# Microsoft Azure App Registration
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here

# Server Configuration
PORT=3000
NODE_ENV=development

# OAuth Scopes (LibreChat manages OAuth flow)
OAUTH_SCOPES=openid profile email offline_access User.Read Mail.Read Calendars.Read Files.Read

# LibreChat Integration Headers
LIBRECHAT_USER_ID_HEADER=x-librechat-user-id
LIBRECHAT_SESSION_HEADER=x-librechat-session
```

## filename: Dockerfile
```dockerfile
FROM node:18-alpine

WORKDIR /app

# Install curl for health checks
RUN apk add --no-cache curl

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production

# Copy source code
COPY . .

# Build the application
RUN npm run build

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/health || exit 1

# Start the server
CMD ["npm", "start"]
```

## filename: docker-compose.yml
```yaml
version: '3.8'

services:
  msgraph-mcp:
    build: .
    ports:
      - "3000:3000"
    environment:
      - NODE_ENV=production
    env_file:
      - .env
    networks:
      - librechat_default
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:3000/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s

networks:
  librechat_default:
    external: true
```

## filename: src/index.ts
```typescript
import express from 'express';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  ErrorCode,
  McpError
} from '@modelcontextprotocol/sdk/types.js';
import dotenv from 'dotenv';
import { GraphTools } from './tools/graphTools.js';
import { logger } from './utils/logger.js';

dotenv.config();

const app = express();
const log = logger('main');

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// CORS headers for LibreChat
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, x-librechat-user-id, x-librechat-session');
  if (req.method === 'OPTIONS') return res.status(200).end();
  next();
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.get('/.well-known/oauth-protected-resource', (req, res) => {
  res.json({
    resource: `http://msgraph-mcp:${process.env.PORT || 3000}`,
    authorization_servers: [
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0`
    ],
    scopes_supported: [
      'User.Read',
      'Mail.Read',
      'Calendars.Read',
      'Files.Read'
    ],
    bearer_methods_supported: ['header']
  });
});

const server = new Server({
  name: 'msgraph-mcp-server',
  version: '1.0.0'
}, {
  capabilities: {
    tools: {},
    resources: {}
  }
});

const graphTools = new GraphTools();

function getAccessToken(req: express.Request): string {
  const header = req.headers['authorization'];
  if (!header || Array.isArray(header)) {
    throw new McpError(ErrorCode.InvalidRequest, 'Missing or invalid Authorization header');
  }
  const parts = header.split(' ');
  if (parts.length !== 2 || parts[0].toLowerCase() !== 'bearer') {
    throw new McpError(ErrorCode.InvalidRequest, 'Authorization header must be a Bearer token');
  }
  return parts[1];
}

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: graphTools.getToolDefinitions()
}));

server.setRequestHandler(CallToolRequestSchema, async (request, { req }) => {
  const token = getAccessToken(req);
  const { name, arguments: args } = request.params;

  log.info(`Calling tool: ${name} with args`, args);

  const graphService = graphTools.createGraphService(token);

  try {
    const result = await graphTools.executeTool(name, args, graphService);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  } catch (err: any) {
    log.error('Tool execution error:', err);
    throw new McpError(ErrorCode.InternalError, `Tool execution failed: ${err.message}`);
  }
});

server.setRequestHandler(ListResourcesRequestSchema, async () => ({
  resources: [
    {
      uri: 'graph://user/profile',
      name: 'User Profile',
      description: 'Current user profile information',
      mimeType: 'application/json'
    },
    {
      uri: 'graph://user/mail',
      name: 'User Mail',
      description: 'User email messages',
      mimeType: 'application/json'
    }
  ]
}));

server.setRequestHandler(ReadResourceRequestSchema, async (request, { req }) => {
  const token = getAccessToken(req);
  const { uri } = request.params;

  const graphService = graphTools.createGraphService(token);

  switch (uri) {
    case 'graph://user/profile':
      return {
        contents: [{
          uri,
          mimeType: 'application/json',
          text: JSON.stringify(await graphService.getUserProfile(), null, 2)
        }]
      };
    case 'graph://user/mail':
      return {
        contents: [{
          uri,
          mimeType: 'application/json',
          text: JSON.stringify(await graphService.getMessages(), null, 2)
        }]
      };
    default:
      throw new McpError(ErrorCode.InvalidRequest, `Unknown resource: ${uri}`);
  }
});

app.all('/mcp', async (req, res) => {
  try {
    const transport = new StreamableHTTPServerTransport(req, res);
    await server.connect(transport, { req });
    log.info('MCP connection handled');
  } catch (error: any) {
    log.error('MCP connection error:', error);
    res.status(400).json({ error: 'MCP_CONNECTION_ERROR', message: error.message });
  }
});

const port = Number(process.env.PORT) || 3000;
app.listen(port, () => {
  log.info(`MS Graph MCP Server running on port ${port}`);
  log.info(`Health check at http://localhost:${port}/health`);
  log.info(`MCP endpoint at http://localhost:${port}/mcp`);
});
```

## filename: src/tools/graphTools.ts
```typescript
import { GraphService } from '../services/graphService.js';
import { logger } from '../utils/logger.js';

const log = logger('graphTools');

export class GraphTools {
  getToolDefinitions() {
    return [
      {
        name: 'get_user_profile',
        description: 'Get the current user profile information from Microsoft Graph',
        inputSchema: {
          type: 'object',
          properties: {},
          required: []
        }
      },
      {
        name: 'get_messages',
        description: 'Get email messages from the user inbox',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Number of messages to retrieve (default: 10, max: 50)',
              default: 10
            }
          },
          required: []
        }
      },
      {
        name: 'get_calendar_events',
        description: 'Get upcoming calendar events for the user',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Number of events to retrieve (default: 10, max: 50)',
              default: 10
            }
          },
          required: []
        }
      },
      {
        name: 'send_email',
        description: 'Send an email message',
        inputSchema: {
          type: 'object',
          properties: {
            subject: {
              type: 'string',
              description: 'Email subject'
            },
            content: {
              type: 'string',
              description: 'Email body content'
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of recipient email addresses'
            }
          },
          required: ['subject', 'content', 'to']
        }
      }
    ];
  }

  createGraphService(token: string) {
    return new GraphService(token);
  }

  async executeTool(toolName: string, args: any, graphService: GraphService): Promise<any> {
    log.info(`Executing tool: ${toolName}`, args);

    switch (toolName) {
      case 'get_user_profile':
        return await graphService.getUserProfile();

      case 'get_messages':
        const messageCount = Math.min(args.top || 10, 50);
        return await graphService.getMessages(messageCount);

      case 'get_calendar_events':
        const eventCount = Math.min(args.top || 10, 50);
        return await graphService.getEvents(eventCount);

      case 'send_email':
        await graphService.sendMail(args.subject, args.content, args.to);
        return { success: true, message: 'Email sent successfully' };

      default:
        throw new Error(`Unknown tool: ${toolName}`);
    }
  }
}
```

## filename: src/services/graphService.ts
```typescript
import axios, { AxiosInstance } from 'axios';
import { GraphUser, GraphMessage, GraphEvent } from '../types/index.js';
import { logger } from '../utils/logger.js';

const log = logger('graphService');

export class GraphService {
  private client: AxiosInstance;

  constructor(accessToken: string) {
    this.client = axios.create({
      baseURL: 'https://graph.microsoft.com/v1.0',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });
  }

  async getUserProfile(): Promise<GraphUser> {
    log.info('Fetching user profile');
    const response = await this.client.get('/me');
    return response.data;
  }

  async getMessages(top: number = 10): Promise<GraphMessage[]> {
    log.info(`Fetching top ${top} messages`);
    const response = await this.client.get(`/me/messages?$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`);
    return response.data.value;
  }

  async getEvents(top: number = 10): Promise<GraphEvent[]> {
    log.info(`Fetching top ${top} calendar events`);
    const response = await this.client.get(`/me/events?$top=${top}&$select=id,subject,start,end,organizer`);
    return response.data.value;
  }

  async sendMail(subject: string, content: string, toRecipients: string[]): Promise<void> {
    log.info(`Sending email with subject: ${subject}`);
    
    const message = {
      subject,
      body: {
        contentType: 'Text',
        content
      },
      toRecipients: toRecipients.map(email => ({
        emailAddress: { address: email }
      }))
    };

    await this.client.post('/me/sendMail', { message });
  }
}
```

## filename: src/utils/logger.ts
```typescript
export const logger = (namespace: string) => {
  const log = (level: string, ...args: any[]) => {
    const timestamp = new Date().toISOString();
    const prefix = `[${timestamp}] [${namespace}] [${level.toUpperCase()}]`;
    console.log(prefix, ...args);
  };

  return {
    info: (...args: any[]) => log('info', ...args),
    warn: (...args: any[]) => log('warn', ...args),
    error: (...args: any[]) => log('error', ...args),
    debug: (...args: any[]) => log('debug', ...args)
  };
};
```

## filename: src/types/index.ts
```typescript
export interface TokenData {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope: string;
}

export interface UserContext {
  userId: string;
  sessionId?: string;
}

export interface GraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
}

export interface GraphMessage {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
  receivedDateTime: string;
  bodyPreview: string;
}

export interface GraphEvent {
  id: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  organizer: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
}
```

## LibreChat Configuration

Add to your `librechat.yaml`:

```yaml
mcpServers:
  msgraph:
    type: streamable-http
    url: http://msgraph-mcp:3000/mcp
    timeout: 30000
    initTimeout: 150000
    serverInstructions: |
      This Microsoft Graph server provides access to:
      - User profile information
      - Email messages (read/send)  
      - Calendar events (read)
      Users must authenticate with Microsoft OAuth2 before using tools.
```

## Azure App Registration Setup

1. Go to Azure Portal > App Registrations
2. Create new app registration
3. Add redirect URI pointing to LibreChat (not MCP server)
4. Generate client secret
5. Add API permissions:
   - User.Read
   - Mail.Read
   - Mail.Send
   - Calendars.Read
6. Grant admin consent