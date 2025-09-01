/* -------------------------------------------------------------------- *
 *  src/MSGraphMCP.ts   –   Pure-Node implementation                    *
 * -------------------------------------------------------------------- */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { MSGraphService } from './MSGraphService.js';
import { Env, MSGraphAuthContext } from '../types.js';
import logger from './lib/logger.js';

/**
 * A self-contained MCP “agent” that is 100% Node-compatible.
 * It does NOT extend the Worker-specific McpAgent – instead we expose our
 * own static `serve()` helper that turns the McpServer into a fetch handler.
 */
export class MSGraphMCP {
  /* ------------------------------------------------------------------ */
  /* construction / lazy service                                         */
  /* ------------------------------------------------------------------ */
  constructor(
    private readonly env: Env,
    private readonly auth: MSGraphAuthContext
  ) {
    logger.info('Initializing MSGraphMCP', { env: { ...env, CLIENT_SECRET: '[REDACTED]', ACCESS_TOKEN: '[REDACTED]' }, auth: { ...auth, accessToken: '[REDACTED]' } });
  }

  #svc?: MSGraphService;
  private get svc(): MSGraphService {
    if (!this.#svc) {
      logger.info('Creating new MSGraphService instance');
      this.#svc = new MSGraphService(this.env, this.auth, {
        tenantId: this.env.TENANT_ID,
        clientId: this.env.CLIENT_ID,
        clientSecret: this.env.CLIENT_SECRET,
        mode: this.env.USE_CLIENT_TOKEN === 'true' ? 'ClientProvidedToken' : this.env.USE_CERTIFICATE === 'true' ? 'Certificate' : this.env.USE_INTERACTIVE === 'true' ? 'Interactive' : 'ClientCredentials',
        redirectUri: this.env.REDIRECT_URI,
        certificatePath: this.env.CERTIFICATE_PATH,
        certificatePassword: this.env.CERTIFICATE_PASSWORD,
      } as any);
    } else {
      logger.info('Reusing existing MSGraphService instance');
    }
    return this.#svc;
  }

  /* ------------------------------------------------------------------ */
  /* helpers                                                             */
  /* ------------------------------------------------------------------ */
  private fmt(label: string, data: unknown) {
    logger.info('Formatting response', { label });
    return {
      content: [
        {
          type: 'text' as const,
          text: `Success! ${label}\n\nResult:\n${JSON.stringify(data, null, 2)}`,
        },
      ],
    };
  }

  /* ------------------------------------------------------------------ */
  /* tool definitions                                                    */
  /* ------------------------------------------------------------------ */
  get server() {
    logger.info('Creating new McpServer instance');
    const server = new McpServer({
      name: 'Microsoft Graph Service',
      version: '1.0.0',
    });

    /* ---------- 1. universal Graph/Azure request -------------------- */
    server.tool(
      'microsoft-graph-api',
      'Versatile Graph / ARM request helper.',
      {
        apiType: z.enum(['graph', 'azure']).optional().default('graph').describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management. Defaults to 'graph'."),
        path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
        method: z.enum(['get', 'post', 'put', 'patch', 'delete']).describe("HTTP method to use"),
        apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
        subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
        queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
        body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
        graphApiVersion: z.enum(['v1.0', 'beta']).optional().default('v1.0').describe("Microsoft Graph API version to use (default: v1.0)"),
        fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
        consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
      },
      async (p) => {
        logger.info('Executing microsoft-graph-api tool', { params: p });
        try {
          if (p.apiType === 'azure') {
            logger.info('Making Azure Resource Management request', { path: p.path, method: p.method });
            const res = await this.svc.azureRequest(
              p.path,
              p.method,
              p.body,
              p.queryParams,
              p.apiVersion,
              p.subscriptionId,
              p.fetchAll
            );
            logger.info('Azure request successful', { path: p.path, method: p.method });
            return this.fmt(`Azure ${p.method.toUpperCase()} ${p.path}`, res);
          }

          logger.info('Making Microsoft Graph request', { path: p.path, method: p.method, graphApiVersion: p.graphApiVersion });
          const res = await this.svc.genericGraphRequest(
            p.path,
            p.method,
            p.body,
            p.queryParams,
            p.graphApiVersion,
            p.fetchAll,
            p.consistencyLevel
          );
          logger.info('Graph request successful', { path: p.path, method: p.method });
          return this.fmt(`Graph ${p.method.toUpperCase()} ${p.path}`, res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('microsoft-graph-api tool error', { msg, params: p });
          throw new Error(msg);
        }
      }
    );

    /* ---------- 2. convenience tools -------------------------------- */
    server.tool(
      'getCurrentUserProfile',
      "Get the current user's Microsoft Graph profile",
      {},
      async () => {
        logger.info('Executing getCurrentUserProfile tool');
        const res = await this.svc.getCurrentUserProfile();
        logger.info('getCurrentUserProfile successful');
        return this.fmt('User profile retrieved', res);
      }
    );

    server.tool(
      'getUsers',
      'Get users from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) => {
        logger.info('Executing getUsers tool', { queryParams, fetchAll });
        const res = await this.svc.getUsers(queryParams, fetchAll);
        logger.info('getUsers successful');
        return this.fmt('Users retrieved', res);
      }
    );

    server.tool(
      'getGroups',
      'Get groups from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) => {
        logger.info('Executing getGroups tool', { queryParams, fetchAll });
        const res = await this.svc.getGroups(queryParams, fetchAll);
        logger.info('getGroups successful');
        return this.fmt('Groups retrieved', res);
      }
    );

    server.tool(
      'getApplications',
      'Get applications from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) => {
        logger.info('Executing getApplications tool', { queryParams, fetchAll });
        const res = await this.svc.getApplications(queryParams, fetchAll);
        logger.info('getApplications successful');
        return this.fmt('Applications retrieved', res);
      }
    );

    /* ---------- 3. Outlook draft ------------------------------------ */
    server.tool(
      'createDraftEmail',
      'Create an Outlook draft (saved to Drafts)',
      {
        subject: z.string().optional(),
        body: z.string().optional(),
        contentType: z.enum(['Text', 'HTML']).optional().default('Text'),
        toRecipients: z.array(z.string()).optional(),
        ccRecipients: z.array(z.string()).optional(),
        bccRecipients: z.array(z.string()).optional(),
      },
      async ({ subject, body, contentType, toRecipients, ccRecipients, bccRecipients }) => {
        logger.info('Executing createDraftEmail tool', { subject, contentType, toRecipients, ccRecipients, bccRecipients });
        try {
          const msg = {
            subject: subject ?? '',
            body: { contentType: contentType ?? 'Text', content: body ?? '' },
            toRecipients: (toRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
            ccRecipients: (ccRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
            bccRecipients: (bccRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
          };
          logger.info('Sending draft email request to MS Graph');
          const res = await this.svc.genericGraphRequest('/me/messages', 'post', msg);
          logger.info('createDraftEmail successful', { response: res });
          return this.fmt('Draft created', res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('createDraftEmail error', { msg, subject, contentType });
          throw new Error(msg);
        }
      }
    );

    /* ---------- Additional convenience tools: Calendar examples ----- */
    server.tool(
      'getUpcomingEvents',
      "Get upcoming calendar events for the current user from Microsoft Graph",
      {
        numberOfEvents: z.number().optional().default(10).describe("Number of events to retrieve. Default: 10"),
        startDateTime: z.string().optional().describe("Start date-time in ISO format to filter events from. Default: current time"),
      },
      async ({ numberOfEvents, startDateTime }) => {
        logger.info('Executing getUpcomingEvents tool', { numberOfEvents, startDateTime });
        try {
          const queryParams = {
            '$top': numberOfEvents.toString(),
            '$orderby': 'start/dateTime',
            '$filter': `start/dateTime ge '${startDateTime || new Date().toISOString()}'`,
          };
          logger.info('Sending get events request to MS Graph', { queryParams });
          const res = await this.svc.genericGraphRequest('/me/events', 'get', null, queryParams, 'v1.0', true);
          logger.info('getUpcomingEvents successful', { response: res });
          return this.fmt('Upcoming events retrieved', res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('getUpcomingEvents tool error', { msg, numberOfEvents, startDateTime });
          throw new Error(msg);
        }
      }
    );

    server.tool(
      'createCalendarEvent',
      "Create a calendar event (meeting) with other people in Microsoft Graph",
      {
        subject: z.string().describe("Subject of the event"),
        startTime: z.string().describe("Start time in ISO format (e.g., '2025-09-02T10:00:00Z')"),
        endTime: z.string().describe("End time in ISO format (e.g., '2025-09-02T11:00:00Z')"),
        attendees: z.array(z.string()).describe("Array of email addresses of attendees"),
        body: z.string().optional().describe("Body content of the event"),
        location: z.string().optional().describe("Location of the event"),
        isOnlineMeeting: z.boolean().optional().default(false).describe("Whether it's an online meeting (e.g., Teams)"),
      },
      async (p) => {
        logger.info('Executing createCalendarEvent tool', { params: p });
        try {
          const event = {
            subject: p.subject,
            start: { dateTime: p.startTime, timeZone: 'UTC' },
            end: { dateTime: p.endTime, timeZone: 'UTC' },
            body: { contentType: 'Text', content: p.body || '' },
            location: { displayName: p.location || '' },
            attendees: p.attendees.map(email => ({ emailAddress: { address: email }, type: 'required' })),
            isOnlineMeeting: p.isOnlineMeeting,
          };
          logger.info('Sending create event request to MS Graph', { event });
          const res = await this.svc.genericGraphRequest('/me/events', 'post', event);
          logger.info('createCalendarEvent successful', { response: res });
          return this.fmt('Calendar event created', res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('createCalendarEvent tool error', { msg, params: p });
          throw new Error(msg);
        }
      }
    );

    logger.info('McpServer configuration completed', { tools: server });
    return server;
  }

  /* ------------------------------------------------------------------ */
  /*  plain-Node helper – turns the server into a fetch handler         */
  /* ------------------------------------------------------------------ */
  static serve() {
    logger.info('Setting up MSGraphMCP serve handler');
    const fetch = async (req: Request): Promise<Response> => {
      logger.info('Received request to MCP server', { method: req.method, url: req.url });

      /* -------- auth extraction ------------------------------------ */
      const authHeader = req.headers.get('Authorization');
      logger.info('Extracting authentication headers', { authHeader: authHeader ? 'Bearer [REDACTED]' : null });
      let auth: MSGraphAuthContext = { accessToken: '' };

      if (authHeader?.startsWith('Bearer ')) {
        auth = {
          accessToken: authHeader.slice(7),
          refreshToken: req.headers.get('x-refresh-token') ?? undefined,
        };
        logger.info('Authentication context extracted', { accessToken: '[REDACTED]', refreshToken: auth.refreshToken ? '[REDACTED]' : undefined });
      } else {
        logger.warn('No valid Bearer token provided in request');
      }

      /* -------- env from process.env ------------------------------- */
      const env: Env = {
        TENANT_ID: process.env.TENANT_ID,
        CLIENT_ID: process.env.CLIENT_ID,
        CLIENT_SECRET: process.env.CLIENT_SECRET,
        ACCESS_TOKEN: process.env.ACCESS_TOKEN,
        REDIRECT_URI: process.env.REDIRECT_URI || 'http://mcp-server:3001/authorize',
        CERTIFICATE_PATH: process.env.CERTIFICATE_PATH,
        CERTIFICATE_PASSWORD: process.env.CERTIFICATE_PASSWORD,
        MS_GRAPH_CLIENT_ID: process.env.MS_GRAPH_CLIENT_ID,
        OAUTH_SCOPES: process.env.OAUTH_SCOPES,
        USE_GRAPH_BETA: process.env.USE_GRAPH_BETA,
        USE_INTERACTIVE: process.env.USE_INTERACTIVE,
        USE_CLIENT_TOKEN: process.env.USE_CLIENT_TOKEN,
        USE_CERTIFICATE: process.env.USE_CERTIFICATE,
      };
      logger.info('Environment variables loaded', { env: { ...env, CLIENT_SECRET: '[REDACTED]', ACCESS_TOKEN: '[REDACTED]' } });

      /* -------- accept only POST with JSON-RPC --------------------- */
      if (req.method !== 'POST') {
        logger.warn('Invalid request method', { method: req.method });
        return new Response('Only POST is allowed', { status: 405 });
      }

      let payload: unknown;
      try {
        payload = await req.json();
        logger.info('Parsed JSON-RPC payload', { payload });
      } catch {
        logger.error('Failed to parse JSON-RPC payload');
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            error: { code: -32700, message: 'Parse error: invalid JSON' },
            id: null,
          }),
          { status: 400, headers: { 'Content-Type': 'application/json' } }
        );
      }

      /* -------- validate JSON-RPC payload ------------------------- */
      if (!payload || typeof payload !== 'object' || !('jsonrpc' in payload) || !('method' in payload)) {
        logger.error('Invalid JSON-RPC payload structure', { payload });
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            error: { code: -32600, message: 'Invalid JSON-RPC request' },
            id: null,
          }),
          { status: 400, headers: { 'Content-Type': 'application/json' } }
        );
      }

      /* -------- delegate to the MCP server ------------------------- */
      logger.info('Creating MSGraphMCP instance for request handling');
      const mcp = new MSGraphMCP(env, auth);
      const rpcServer = mcp.server;

      try {
        logger.info('Calling serveRequest with payload', { method: (payload as any).method, id: (payload as any).id });
        /* Temporary workaround: Use type assertion due to unknown SDK method */
        const result = await (rpcServer as any).serveRequest(payload);
        logger.info('serveRequest executed successfully', { result });
        return new Response(JSON.stringify(result), {
          headers: { 'Content-Type': 'application/json' },
        });
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        logger.error('MCP execution error', { error: msg, payload });
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            error: {
              code: -32000,
              message: msg,
            },
            id: typeof payload === 'object' && payload ? (payload as any).id ?? null : null,
          }),
          { status: 500, headers: { 'Content-Type': 'application/json' } }
        );
      }
    };

    /* Hono’s .mount() expects { fetch } */
    logger.info('Returning fetch handler for MSGraphMCP');
    return { fetch };
  }
}