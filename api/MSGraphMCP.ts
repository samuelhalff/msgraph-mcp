import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { MSGraphService } from "./MSGraphService.js"; // AuthManager, etc. are no longer needed here
import { MSGraphAuthContext, Env } from "../types";
import logger from "./lib/logger.js";

export class MSGraphMCP {
  private env: Env;
  private authContext: MSGraphAuthContext;
  private _msGraphService: MSGraphService | null = null;

  constructor(env: Env, authContext: MSGraphAuthContext) {
    this.env = env;
    this.authContext = authContext;
  }

  private get msGraphServiceInstance(): MSGraphService {
    if (!this._msGraphService) {
      const authConfig = {
        tenantId: this.env.TENANT_ID,
        clientId: this.env.CLIENT_ID,
        clientSecret: this.env.CLIENT_SECRET,
      } as any;

      this._msGraphService = new MSGraphService(
        this.env,
        this.authContext,
        authConfig
      );
    }
    return this._msGraphService;
  }

  private formatResponse = (
    description: string,
    data: unknown
  ): {
    content: Array<{ type: "text"; text: string }>;
  } => {
    return {
      content: [
        {
          type: "text",
          text: `Success! ${description}\n\nResult:\n${JSON.stringify(
            data,
            null,
            2
          )}`,
        },
      ],
    };
  };

  // This is the definition of your server and all its tools.
  // LibreChat calls this getter to understand what your agent can do.
  get server() {
    const server = new McpServer({
      name: "Microsoft Graph Service",
      version: "1.0.0",
    });

    // ========================================================================
    // YOUR COMPLETE TOOL DEFINITIONS ARE PRESERVED HERE
    // ========================================================================

    // Main Microsoft Graph API tool
    server.tool(
      "microsoft-graph-api",
      "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. IMPORTANT: For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: \"eventual\"'.",
      {
        apiType: z
          .enum(["graph", "azure"])
          .optional()
          .default("graph")
          .describe(
            "Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management. Defaults to 'graph'."
          ),
        path: z
          .string()
          .describe(
            "The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"
          ),
        method: z
          .enum(["get", "post", "put", "patch", "delete"])
          .describe("HTTP method to use"),
        apiVersion: z
          .string()
          .optional()
          .describe(
            "Azure Resource Management API version (required for apiType Azure)"
          ),
        subscriptionId: z
          .string()
          .optional()
          .describe("Azure Subscription ID (for Azure Resource Management)."),
        queryParams: z
          .record(z.string())
          .optional()
          .describe("Query parameters for the request"),
        body: z
          .record(z.string(), z.any())
          .optional()
          .describe("The request body (for POST, PUT, PATCH)"),
        graphApiVersion: z
          .enum(["v1.0", "beta"])
          .optional()
          .default("v1.0")
          .describe("Microsoft Graph API version to use (default: v1.0)"),
        fetchAll: z
          .boolean()
          .optional()
          .default(false)
          .describe(
            "Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."
          ),
        consistencyLevel: z
          .string()
          .optional()
          .describe(
            "Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."
          ),
      },
      async (params) => {
        try {
          // Accessing the service instance here will trigger the auth check.
          const responseData =
            await this.msGraphServiceInstance.genericGraphRequest(
              params.path,
              params.method,
              params.body,
              params.queryParams,
              params.graphApiVersion,
              params.fetchAll,
              params.consistencyLevel
            );

          let resultText = `Result for ${params.apiType} API (${params.graphApiVersion}) - ${params.method} ${params.path}:\n\n`;
          resultText += JSON.stringify(responseData, null, 2);
          // ... (your logic for adding the 'Note: More results' text) ...

          return { content: [{ type: "text" as const, text: resultText }] };
        } catch (err: unknown) {
          const message = err instanceof Error ? err.message : String(err);
          logger.error("Error in microsoft-graph-api tool", { message });
          throw new Error(message); // Re-throw so the SDK can format the error.
        }
      }
    );

    // Your other specific tools
    server.tool(
      "getCurrentUserProfile",
      "Get the current user's Microsoft Graph profile",
      {},
      async () => {
        const profile =
          await this.msGraphServiceInstance.getCurrentUserProfile();
        return this.formatResponse("User profile retrieved", profile);
      }
    );

    server.tool(
      "getUsers",
      "Get users from Microsoft Graph",
      {
        queryParams: z
          .record(z.string())
          .optional()
          .describe("Query parameters for filtering users"),
        fetchAll: z
          .boolean()
          .optional()
          .default(false)
          .describe("Set to true to fetch all users"),
      },
      async ({ queryParams, fetchAll }) => {
        const users = await this.msGraphServiceInstance.getUsers(
          queryParams,
          fetchAll
        );
        return this.formatResponse("Users retrieved", users);
      }
    );

    server.tool(
      "getGroups",
      "Get groups from Microsoft Graph",
      {
        queryParams: z
          .record(z.string())
          .optional()
          .describe("Query parameters for filtering groups"),
        fetchAll: z
          .boolean()
          .optional()
          .default(false)
          .describe("Set to true to fetch all groups"),
      },
      async ({ queryParams, fetchAll }) => {
        const groups = await this.msGraphServiceInstance.getGroups(
          queryParams,
          fetchAll
        );
        return this.formatResponse("Groups retrieved", groups);
      }
    );

    server.tool(
      "createDraftEmail",
      "Create an email draft in Outlook",
      {
        subject: z.string().optional(),
        body: z.string().optional(),
        // ... your full Zod schema for this tool
      },
      async (params) => {
        const message = {
          subject: params.subject,
          body: {
            content: params.body,
            contentType: "Text",
          } /* build full message object */,
        };
        const response = await this.msGraphServiceInstance.genericGraphRequest(
          "/me/messages",
          "post",
          message
        );
        return this.formatResponse("Draft created", response);
      }
    );

    // ... and so on for ALL your other tools.

    // NOTE: As this architecture is stateless per-request, tools like `set-access-token`
    // and `get-auth-status` are no longer meaningful and can be removed. LibreChat's
    // framework is now responsible for the token.

    return server;
  }

  // This is the main entry point that LibreChat will call for every request.
  static serve() {
    return {
      fetch: async (request: Request): Promise<Response> => {
        // 1.  Extract bearer / refresh tokens that LibreChat sends
        const authHeader = request.headers.get('Authorization') ?? '';
        const authContext: MSGraphAuthContext = {
          accessToken: authHeader.startsWith('Bearer ')
            ? authHeader.slice(7)
            : '',
          refreshToken: request.headers.get('x-refresh-token') ?? undefined,
        };

        // 2.  Load env variables that your service needs
        const env: Env = {
          TENANT_ID:     process.env.TENANT_ID     || '',
          CLIENT_ID:     process.env.CLIENT_ID     || '',
          CLIENT_SECRET: process.env.CLIENT_SECRET || '',
        };

        try {
          // 3.  Create a one-off MSGraphMCP instance for this request
          const mcp = new MSGraphMCP(env, authContext);

          // 4.  Obtain the SDK’s JSON-RPC fetch handler
          //     (in SDK >= 0.3 you call .serve(); older versions already expose {fetch})
          const rpcServer = (mcp.server as any).serve
            ? (mcp.server as any).serve()   // returns { fetch }
            : (mcp.server as any);          // already { fetch }

          // 5.  Delegate the request to the real handler
          return await rpcServer.fetch(request);

        } catch (error: unknown) {
          const message =
            error instanceof Error ? error.message : String(error);
          logger.error('Fatal error in the fetch handler', { message });

          // What you previously returned on error:
          return new Response(
            JSON.stringify({
              jsonrpc: '2.0',
              error: { code: -32603, message: 'Internal Server Error' },
              id: null,
            }),
            { status: 500, headers: { 'Content-Type': 'application/json' } },
          );
        }
      },
    };
  }
}
