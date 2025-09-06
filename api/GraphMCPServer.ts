import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  InitializeRequestSchema,
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ToolListChangedNotification,
  JSONRPCNotification,
  Notification,
} from "@modelcontextprotocol/sdk/types.js";
import type { Request, Response } from "express";
import { randomUUID } from "crypto";
import logger from "./lib/logger.js";
import { MSGraphMCP } from "./MSGraphMCP.js";
import type { Env, MSGraphAuthContext } from "../types.js";

const SESSION_ID_HEADER_NAME = "mcp-session-id";

export class GraphMCPServer {
  server: Server;
  transports: Record<string, StreamableHTTPServerTransport> = {};
  // Per-session auth context (e.g., bearer token)
  private sessionAuth: Record<string, MSGraphAuthContext> = {};
  private toolInterval: NodeJS.Timeout | undefined;

  constructor(server: Server) {
    this.server = server;
    this.setupHandlers();
  }

  // Allow external router to update per-session auth context
  public setSessionAuth(sessionId: string, accessToken: string) {
    this.sessionAuth[sessionId] = { accessToken };
    logger.info("Session auth updated", { sessionId, hasToken: !!accessToken });
  }

  // Wire MCP request handlers to our MSGraphMCP tools
  private setupHandlers() {
    // initialize
    this.server.setRequestHandler(
      InitializeRequestSchema,
      async (_req, extra) => {
        // Send initialized notification after responding
        setTimeout(() => {
          const transport = (extra as any)?.transport as
            | StreamableHTTPServerTransport
            | undefined;
          if (transport) {
            const notification: Notification = {
              method: "notifications/initialized",
            } as any;
            this.sendNotification(transport, notification).catch((e) =>
              logger.warn("Failed to send initialized notification", {
                error: String(e),
              })
            );
          }
        }, 0);

        return {
          protocolVersion: "2024-11-05",
          capabilities: { tools: {}, logging: {} },
          serverInfo: { name: "Microsoft Graph MCP Server", version: "1.0.0" },
        };
      }
    );

    // tools/list
    this.server.setRequestHandler(
      ListToolsRequestSchema,
      async (_req, extra) => {
        const { env } = this.getEnvAndAuthFromExtra(extra);
        const mcp = new MSGraphMCP(env, { accessToken: "" });
        return { tools: mcp.getAvailableTools() };
      }
    );

    // tools/call
    this.server.setRequestHandler(
      CallToolRequestSchema,
      async (request, extra) => {
        const { env, auth } = this.getEnvAndAuthFromExtra(extra);
        const mcp = new MSGraphMCP(env, auth);
        // Ensure tools are registered and handlers populated
        void mcp.server;
        const name = request.params.name;
        const args = request.params.arguments || {};
        logger.info("tools/call invoked", {
          name,
          hasToken: !!auth.accessToken,
        });
        return await mcp.callTool(name, args as Record<string, unknown>);
      }
    );

    // Optional: periodically notify tool list changes (no-op but example parity)
    this.toolInterval = setInterval(() => {
      const notification: ToolListChangedNotification = {
        method: "notifications/tools/list_changed",
      };
      for (const transport of Object.values(this.transports)) {
        this.sendNotification(transport, notification).catch((e) =>
          logger.warn("Failed to send tool list notification", {
            error: String(e),
          })
        );
      }
    }, 5 * 60 * 1000);
  }

  private getEnv(): Env {
    // Build env from process.env; ACCESS_TOKEN is set per-session via sessionAuth
    return {
      TENANT_ID: process.env.TENANT_ID,
      CLIENT_ID: process.env.CLIENT_ID,
      CLIENT_SECRET: process.env.CLIENT_SECRET,
      REDIRECT_URI: process.env.REDIRECT_URI,
      AUTH_BASE_URL: process.env.AUTH_BASE_URL,
      GRAPH_BASE_URL: process.env.GRAPH_BASE_URL,
      USE_CLIENT_TOKEN: process.env.USE_CLIENT_TOKEN,
      USE_CERTIFICATE: process.env.USE_CERTIFICATE,
      USE_INTERACTIVE: process.env.USE_INTERACTIVE,
      USE_GRAPH_BETA: process.env.USE_GRAPH_BETA,
      CERTIFICATE_PATH: process.env.CERTIFICATE_PATH,
      CERTIFICATE_PASSWORD: process.env.CERTIFICATE_PASSWORD,
      SCOPES: process.env.SCOPES,
      PORT: process.env.PORT,
    } as Env;
  }

  private getEnvAndAuthFromExtra(extra: unknown): {
    env: Env;
    auth: MSGraphAuthContext;
  } {
    const env = this.getEnv();
    const sessionId = this.getSessionIdFromExtra(extra);
    const auth = (sessionId && this.sessionAuth[sessionId]) || {
      accessToken: "",
    };
    // Inject ACCESS_TOKEN for client-provided token mode
    const envWithToken = { ...env, ACCESS_TOKEN: auth.accessToken } as Env;
    return { env: envWithToken, auth };
  }

  private getSessionIdFromExtra(extra: any): string | undefined {
    try {
      const sid = extra?.transport?.sessionId as string | undefined;
      return sid;
    } catch {
      return undefined;
    }
  }

  async handleGetRequest(req: Request, res: Response) {
    // Follow example: only allow GET for SSE with valid session id
    const sessionId = req.headers[SESSION_ID_HEADER_NAME] as string | undefined;
    const userAgent = req.headers['user-agent'];
    const acceptHeader = req.headers['accept'];
    const origin = req.headers['origin'];
    
    logger.info("SSE GET request received", {
      sessionId: sessionId || "none",
      userAgent,
      acceptHeader,
      origin,
      hasTransport: sessionId ? !!this.transports[sessionId] : false,
      activeSessions: Object.keys(this.transports).length,
      ip: req.ip || req.connection?.remoteAddress,
      url: req.url,
      headers: req.headers
    });

    if (!sessionId || !this.transports[sessionId]) {
      logger.warn("SSE GET rejected - invalid session", {
        sessionId: sessionId || "missing",
        availableSessions: Object.keys(this.transports),
        reason: !sessionId ? "no session header" : "session not found"
      });
      res
        .status(200)
        .json({
          jsonrpc: "2.0",
          id: null,
          error: { code: -32601, message: "Method not found" },
        });
      return;
    }

    logger.info(`Establishing SSE stream for session ${sessionId}`, {
      transportExists: !!this.transports[sessionId],
      sessionAuth: !!this.sessionAuth[sessionId],
      clientInfo: { userAgent, origin }
    });
    
    const transport = this.transports[sessionId];
    
    // Add connection lifecycle logging
    res.on('close', () => {
      logger.info(`SSE connection closed for session ${sessionId}`, {
        userAgent,
        reason: 'client_disconnect'
      });
    });

    res.on('error', (error) => {
      logger.error(`SSE connection error for session ${sessionId}`, {
        error: error.message,
        userAgent,
        stack: error.stack
      });
    });

    // Set proxy-safe headers for SSE and disable buffering when possible
    try {
      // Do not override Content-Type; transport will set it to text/event-stream
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('Keep-Alive', 'timeout=600');
      // Disable proxy buffering (nginx)
      res.setHeader('X-Accel-Buffering', 'no');
    } catch {}

    // Heartbeat to keep idle connections alive through proxies (SSE comments are safe)
    const HEARTBEAT_MS = 15000; // 15s
    let heartbeat: NodeJS.Timeout | undefined;
    const startHeartbeat = () => {
      if (heartbeat) return;
      logger.debug(`Starting SSE heartbeat for session ${sessionId}`, {
        intervalMs: HEARTBEAT_MS
      });
      heartbeat = setInterval(() => {
        if (res.writable && !res.writableEnded) {
          try {
            // SSE comment line; ignored by clients but keeps connection active
            res.write(`: ping ${Date.now()}\n\n`);
          } catch (e) {
            logger.debug(`Heartbeat write failed for session ${sessionId}`, {
              error: e instanceof Error ? e.message : String(e)
            });
          }
        }
      }, HEARTBEAT_MS);
    };

    const stopHeartbeat = () => {
      if (heartbeat) {
        clearInterval(heartbeat);
        heartbeat = undefined;
        logger.debug(`Stopped SSE heartbeat for session ${sessionId}`);
      }
    };

    // Ensure we stop heartbeat on connection end/error
    res.once('close', stopHeartbeat);
    res.once('error', stopHeartbeat);

    try {
      // Start heartbeat just before handing over to transport
      startHeartbeat();
      await transport.handleRequest(req as any, res as any);
      logger.info(`SSE transport handling completed for session ${sessionId}`);
    } catch (error) {
      logger.error(`SSE transport handling failed for session ${sessionId}`, {
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        userAgent
      });
      throw error;
    } finally {
      stopHeartbeat();
    }
    return;
  }

  async handlePostRequest(req: Request, res: Response) {
    const sessionId = req.headers[SESSION_ID_HEADER_NAME] as string | undefined;
    const authHeader = req.header("Authorization");
    const bearer = authHeader?.startsWith("Bearer ")
      ? authHeader.slice(7)
      : undefined;
    const acceptHeader = req.header("accept");

  // Enforce auth for tools/call before handing to transport so we can return 401 with header
  // Body may arrive as text/plain from node-fetch defaults; try to parse JSON if needed
    let body: any = req.body;
    if (typeof body === "string") {
      try {
        body = JSON.parse(body);
      } catch {
        // leave as-is; will result in 400 below if not initialize
      }
    }
  const isBatch = Array.isArray(body);
    const bodyArray = isBatch ? (body as any[]) : [body];
  logger.info("MCP POST request received with body:", { bodyArray });
    const isInitEarly = this.isInitializeRequest(body);
    const containsToolsCall = bodyArray.some((r) => r?.method === "tools/call");
    // If there's no session yet and this isn't initialize or tools/call, return method-not-found
    if (!sessionId && !isInitEarly && !containsToolsCall) {
      logger.info("Early method-not-found path", {
        hasSessionId: !!sessionId,
        isInitEarly,
        containsToolsCall,
        method: body?.method,
      });
      const toRpcError = (item: any) => ({
        jsonrpc: "2.0",
        id: item?.id ?? null,
        error: { code: -32601, message: "Method not found" },
      });
      const response = Array.isArray(body)
        ? body.map((i: any) => toRpcError(i))
        : toRpcError(body);
      res.status(200).json(response);
      return;
    }
    // Determine if request contains a known, non-whitelisted tool call
    // Build a temp MSGraphMCP to consult the registered tool list
    const allowlist = new Set<string>(["throttling-stats"]);
    let containsProtectedCall = false;
    try {
      const env = this.getEnv();
      const mcp = new MSGraphMCP(
        { ...env, ACCESS_TOKEN: bearer || "" } as Env,
        {
          accessToken: bearer || "",
        }
      );
      // Access server getter to ensure tools are registered
      void mcp.server;
      for (const r of bodyArray) {
        if (r?.method === "tools/call") {
          const toolName = r?.params?.name as string | undefined;
          if (!toolName) continue;
          const isKnown = mcp.hasTools(toolName);
          const isAllowed = allowlist.has(toolName);
          if (isKnown && !isAllowed) {
            containsProtectedCall = true;
            break;
          }
        }
      }
    } catch (e) {
      // If any error in gating logic, fall back to not gating to allow proper JSON-RPC errors downstream
      containsProtectedCall = false;
      logger.warn("Auth gating check failed, skipping pre-auth", {
        error: e instanceof Error ? e.message : String(e),
      });
    }

    // Debug logging for initialize detection & headers
    try {
      const debugMethod = isBatch
        ? bodyArray.map((r) => r?.method).join(",")
        : body?.method;
      const debugJsonRpc = isBatch
        ? bodyArray.map((r) => r?.jsonrpc).join(",")
        : body?.jsonrpc;
      logger.info("MCP POST received", {
        hasSessionId: !!sessionId,
        methods: debugMethod,
        jsonrpc: debugJsonRpc,
        isBatch,
        containsProtectedCall,
        hasBearer: !!bearer,
        acceptHeader,
      });
    } catch {}

    if (containsProtectedCall && !bearer) {
      const resourceMetadataUrl = this.getResourceMetadataUrl(req);
      res.setHeader(
        "WWW-Authenticate",
        `Bearer realm="MCP", resource_metadata_url="${resourceMetadataUrl}"`
      );
      res.status(401).json({
        jsonrpc: "2.0",
        id: body?.id ?? null,
        error: {
          code: -32002,
          message: "Authentication required for this method",
        },
      });
      return;
    }

    try {
      // Reuse existing session if present
      if (sessionId && this.transports[sessionId]) {
        logger.info(`Reusing existing session ${sessionId}`, {
          hasAuth: !!bearer,
          willUpdateAuth: !!bearer,
          userAgent: req.headers['user-agent']
        });
        if (bearer) {
          this.sessionAuth[sessionId] = { accessToken: bearer };
          logger.info(`Updated auth for existing session ${sessionId}`);
        }
        const transport = this.transports[sessionId];
        
        try {
          await transport.handleRequest(req as any, res as any, body);
          logger.info(`Successfully handled request for session ${sessionId}`);
        } catch (error) {
          logger.error(`Transport request handling failed for session ${sessionId}`, {
            error: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined
          });
          throw error;
        }
        return;
      }

      // Create new transport if this is an initialize request
  const isInit = this.isInitializeRequest(body);
      if (!sessionId && isInit) {
        logger.info(
          "Detected initialize request; creating new streamable transport", {
            userAgent: req.headers['user-agent'],
            origin: req.headers['origin'],
            ip: req.ip || req.connection?.remoteAddress,
            acceptHeader,
            bodyPreview: JSON.stringify(body).substring(0, 200)
          }
        );
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
        });

        logger.info("Connecting server to new transport");
        await this.server.connect(transport);
        
        logger.info("Handling initialize request with new transport");
        await transport.handleRequest(req as any, res as any, body);

        const sid = (transport as any).sessionId as string | undefined;
        if (sid) {
          this.transports[sid] = transport;
          logger.info(`Transport registered with session ID: ${sid}`, {
            totalSessions: Object.keys(this.transports).length,
            hasAuth: !!bearer
          });
          if (bearer) {
            this.sessionAuth[sid] = { accessToken: bearer };
            logger.info(`Auth stored for session ${sid}`);
          }
        } else {
          logger.warn("Transport created but no session ID available");
        }
        return;
      }

      // If we have a JSON-RPC-shaped request without a valid session, return method-not-found
      if (body && (body.method || (Array.isArray(body) && body.length > 0))) {
        const toRpcError = (item: any) => ({
          jsonrpc: "2.0",
          id: item?.id ?? null,
          error: { code: -32601, message: "Method not found" },
        });
        const response = Array.isArray(body)
          ? body.map((i: any) => toRpcError(i))
          : toRpcError(body);
        res.status(200).json(response);
        return;
      }

      res
        .status(400)
        .json(
          this.createErrorResponse("Bad Request: invalid session ID or method.")
        );
      return;
    } catch (error) {
      logger.error("Error handling MCP request", {
        error: error instanceof Error ? error.message : String(error),
      });
      res.status(500).json(this.createErrorResponse("Internal server error."));
      return;
    }
  }

  async cleanup() {
    logger.info("Starting GraphMCPServer cleanup", {
      activeSessions: Object.keys(this.transports).length,
      hasToolInterval: !!this.toolInterval
    });
    
    this.toolInterval && clearInterval(this.toolInterval);
    
    // Clean up all transports
    for (const [sessionId, transport] of Object.entries(this.transports)) {
      try {
        logger.info(`Cleaning up transport for session ${sessionId}`);
        // If transport has cleanup method, call it
        if (typeof (transport as any).close === 'function') {
          await (transport as any).close();
        }
      } catch (error) {
        logger.warn(`Failed to cleanup transport ${sessionId}`, {
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
    
    this.transports = {};
    this.sessionAuth = {};
    
    await this.server.close();
    logger.info("GraphMCPServer cleanup completed");
  }

  private async sendNotification(
    transport: StreamableHTTPServerTransport,
    notification: Notification
  ) {
    const rpcNotification: JSONRPCNotification = {
      ...notification,
      jsonrpc: "2.0",
    } as any;
    
    try {
      logger.debug("Sending notification", {
        method: notification.method,
        sessionId: (transport as any).sessionId,
        notificationType: typeof notification
      });
      await (transport as any).send(rpcNotification);
      logger.debug("Notification sent successfully", {
        method: notification.method
      });
    } catch (error) {
      logger.error("Failed to send notification", {
        method: notification.method,
        error: error instanceof Error ? error.message : String(error),
        sessionId: (transport as any).sessionId
      });
      throw error;
    }
  }

  private createErrorResponse(message: string) {
    return {
      jsonrpc: "2.0",
      error: { code: -32000, message },
      id: randomUUID(),
    };
  }

  private isInitializeRequest(body: any): boolean {
    // Be liberal in detection: only rely on JSON-RPC shape, not SDK schema parsing
    const isInitShape = (data: any) =>
      !!data &&
      typeof data === "object" &&
      data.method === "initialize" &&
      (data.jsonrpc === "2.0" || typeof data.jsonrpc === "string");

    if (Array.isArray(body)) return body.some((r) => isInitShape(r));
    return isInitShape(body);
  }

  private getResourceMetadataUrl(req: Request) {
    const proto = (req.headers["x-forwarded-proto"] as string) || "https";
    const host = req.headers["host"] as string;
    return `${proto}://${host}/.well-known/oauth-protected-resource`;
  }
}
