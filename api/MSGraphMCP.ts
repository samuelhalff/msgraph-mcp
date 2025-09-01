/* -------------------------------------------------------------------- *
 *  src/MSGraphMCP.ts   –   Pure-Node implementation                    *
 * -------------------------------------------------------------------- */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';

import { MSGraphService } from './MSGraphService.js';
import { MSGraphAuthContext, Env } from '../types.js';
import logger from './lib/logger.js';

/**
 * A self-contained MCP “agent” that is 100 % Node-compatible.
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
  ) {}

  #svc?: MSGraphService;
  private get svc(): MSGraphService {
    if (!this.#svc) {
      this.#svc = new MSGraphService(this.env, this.auth, {
        tenantId: this.env.TENANT_ID,
        clientId: this.env.CLIENT_ID,
        clientSecret: this.env.CLIENT_SECRET,
      } as any);
    }
    return this.#svc;
  }

  /* ------------------------------------------------------------------ */
  /* helpers                                                             */
  /* ------------------------------------------------------------------ */
  private fmt(label: string, data: unknown) {
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
    const server = new McpServer({
      name: 'Microsoft Graph Service',
      version: '1.0.0',
    });

    /* ---------- 1. universal Graph/Azure request -------------------- */
    server.tool(
      'microsoft-graph-api',
      'Versatile Graph / ARM request helper.',
      {
        apiType: z.enum(['graph', 'azure']).optional().default('graph'),
        path: z.string(),
        method: z.enum(['get', 'post', 'put', 'patch', 'delete']),
        apiVersion: z.string().optional(),
        subscriptionId: z.string().optional(),
        queryParams: z.record(z.string()).optional(),
        body: z.record(z.string(), z.any()).optional(),
        graphApiVersion: z.enum(['v1.0', 'beta']).optional().default('v1.0'),
        fetchAll: z.boolean().optional().default(false),
        consistencyLevel: z.string().optional(),
      },
      async (p) => {
        try {
          if (p.apiType === 'azure') {
            const res = await this.svc.azureRequest(
              p.path,
              p.method,
              p.body,
              p.queryParams,
              p.apiVersion,
              p.subscriptionId,
              p.fetchAll
            );
            return this.fmt(`Azure ${p.method.toUpperCase()} ${p.path}`, res);
          }

          const res = await this.svc.genericGraphRequest(
            p.path,
            p.method,
            p.body,
            p.queryParams,
            p.graphApiVersion,
            p.fetchAll,
            p.consistencyLevel
          );
          return this.fmt(`Graph ${p.method.toUpperCase()} ${p.path}`, res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('microsoft-graph-api tool error', { msg });
          throw new Error(msg);
        }
      }
    );

    /* ---------- 2. convenience tools -------------------------------- */
    server.tool(
      'getCurrentUserProfile',
      "Get the current user's Microsoft Graph profile",
      {},
      async () => this.fmt('User profile retrieved', await this.svc.getCurrentUserProfile())
    );

    server.tool(
      'getUsers',
      'Get users from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) =>
        this.fmt('Users retrieved', await this.svc.getUsers(queryParams, fetchAll))
    );

    server.tool(
      'getGroups',
      'Get groups from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) =>
        this.fmt('Groups retrieved', await this.svc.getGroups(queryParams, fetchAll))
    );

    server.tool(
      'getApplications',
      'Get applications from Microsoft Graph',
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) =>
        this.fmt('Applications retrieved', await this.svc.getApplications(queryParams, fetchAll))
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
        try {
          const msg = {
            subject: subject ?? '',
            body: { contentType: contentType ?? 'Text', content: body ?? '' },
            toRecipients: (toRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
            ccRecipients: (ccRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
            bccRecipients: (bccRecipients ?? []).map((a) => ({ emailAddress: { address: a } })),
          };

          const res = await this.svc.genericGraphRequest('/me/messages', 'post', msg);
          return this.fmt('Draft created', res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error('createDraftEmail error', { msg });
          throw new Error(msg);
        }
      }
    );

    return server;
  }

  /* ------------------------------------------------------------------ */
  /*  plain-Node helper – turns the server into a fetch handler         */
  /* ------------------------------------------------------------------ */
static serve() {
    const fetch = async (req: Request): Promise<Response> => {
      /* -------- auth extraction ------------------------------------ */
      const authHeader = req.headers.get('Authorization');
      let auth: MSGraphAuthContext = { accessToken: '' };

      if (authHeader?.startsWith('Bearer ')) {
        auth = {
          accessToken: authHeader.slice(7),
          refreshToken: req.headers.get('x-refresh-token') ?? undefined,
        };
      }

      /* -------- env from process.env ------------------------------- */
      const env: Env = {
        TENANT_ID: process.env.TENANT_ID,
        CLIENT_ID: process.env.CLIENT_ID,
        CLIENT_SECRET: process.env.CLIENT_SECRET,
        ACCESS_TOKEN: process.env.ACCESS_TOKEN,
        REDIRECT_URI: process.env.REDIRECT_URI,
        CERTIFICATE_PATH: process.env.CERTIFICATE_PATH,
        CERTIFICATE_PASSWORD: process.env.CERTIFICATE_PASSWORD,
        MS_GRAPH_CLIENT_ID: process.env.MS_GRAPH_CLIENT_ID,
        OAUTH_SCOPES: process.env.OAUTH_SCOPES,
        USE_GRAPH_BETA: process.env.USE_GRAPH_BETA,
        USE_INTERACTIVE: process.env.USE_INTERACTIVE,
        USE_CLIENT_TOKEN: process.env.USE_CLIENT_TOKEN,
        USE_CERTIFICATE: process.env.USE_CERTIFICATE,
      };

      /* -------- accept only POST with JSON-RPC --------------------- */
      if (req.method !== 'POST') {
        return new Response('Only POST is allowed', { status: 405 });
      }

      let payload: unknown;
      try {
        payload = await req.json();
      } catch {
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            error: { code: -32700, message: 'Parse error: invalid JSON' },
            id: null,
          }),
          { status: 400, headers: { 'Content-Type': 'application/json' } }
        );
      }

      /* -------- delegate to the MCP server ------------------------- */
      const mcp = new MSGraphMCP(env, auth);
      const rpcServer = mcp.server;

      try {
        /* most SDK builds expose one of these names */
        const handler =
          (rpcServer as any).handle ??
          (rpcServer as any).dispatch ??
          (rpcServer as any).process ??
          (rpcServer as any).dispatchRequest;

        if (typeof handler !== 'function') {
          throw new Error(
            'Unsupported SDK version: no handle/dispatch/process function found'
          );
        }

        const result = await handler.call(rpcServer, payload);

        return new Response(JSON.stringify(result), {
          headers: { 'Content-Type': 'application/json' },
        });
      } catch (err) {
        logger.error('MCP execution error', err);
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            error: {
              code: -32000,
              message: err instanceof Error ? err.message : String(err),
            },
            // echo id if available
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            id: typeof payload === 'object' && payload ? payload.id ?? null : null,
          }),
          { status: 500, headers: { 'Content-Type': 'application/json' } }
        );
      }
    };

    /* Hono’s .mount() expects { fetch } */
    return { fetch };
  }
}