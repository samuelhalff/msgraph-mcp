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
    if (!sessionId || !this.transports[sessionId]) {
      res
        .status(200)
        .json({
          jsonrpc: "2.0",
          id: null,
          error: { code: -32601, message: "Method not found" },
        });
      return;
    }

    logger.info(`Establishing SSE stream for session ${sessionId}`);
    const transport = this.transports[sessionId];
    await transport.handleRequest(req as any, res as any);
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
        if (bearer) {
          this.sessionAuth[sessionId] = { accessToken: bearer };
        }
        const transport = this.transports[sessionId];
        await transport.handleRequest(req as any, res as any, body);
        return;
      }

      // Create new transport if this is an initialize request
  const isInit = this.isInitializeRequest(body);
      if (!sessionId && isInit) {
        logger.info(
          "Detected initialize request; creating new streamable transport"
        );
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
        });

        await this.server.connect(transport);
        await transport.handleRequest(req as any, res as any, body);

        const sid = (transport as any).sessionId as string | undefined;
        if (sid) {
          this.transports[sid] = transport;
          if (bearer) {
            this.sessionAuth[sid] = { accessToken: bearer };
          }
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
    this.toolInterval && clearInterval(this.toolInterval);
    await this.server.close();
  }

  private async sendNotification(
    transport: StreamableHTTPServerTransport,
    notification: Notification
  ) {
    const rpcNotification: JSONRPCNotification = {
      ...notification,
      jsonrpc: "2.0",
    } as any;
    await (transport as any).send(rpcNotification);
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
