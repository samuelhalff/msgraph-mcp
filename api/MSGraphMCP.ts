// src/MSGraphMCP.ts
import { McpAgent } from "agents/mcp"; // <-- identical import path used by SpotifyMCP
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import { MSGraphService } from "./MSGraphService.js";
import { MSGraphAuthContext, Env } from "../types";
import logger from "./lib/logger.js";

// NB: Generic signature = <Env-type,  Custom-state-type, Auth-context-type >
export class MSGraphMCP extends McpAgent<Env, unknown, MSGraphAuthContext> {
  /*******************
   *  Optional init  *
   *******************/
  async init() {
    /*  Put one-time initialisation here if you want (not required). */
  }

  /******************************************************
   *  Lazy-instantiated wrapper around MSGraphService   *
   ******************************************************/
  #service?: MSGraphService; // private field (Node ≥ 16 / TS 4.4+)

  private get graphService(): MSGraphService {
    if (!this.#service) {
      const env = (this as any).env ?? (process.env as unknown as Env);

      const authConfig = {
        tenantId: env.TENANT_ID,
        clientId: env.CLIENT_ID,
        clientSecret: env.CLIENT_SECRET,
      } as any;

      /*  In a McpAgent, the OAuth tokens arrive in  this.props  */
      this.#service = new MSGraphService(
        env,
        (this as any).props, // MSGraphAuthContext { accessToken, refreshToken }
        authConfig
      );
    }
    return this.#service;
  }

  /***********************
   *  Helper formatter   *
   ***********************/
  private formatResponse(description: string, data: unknown) {
    return {
      content: [
        {
          type: "text" as const,
          text: `Success! ${description}\n\nResult:\n${JSON.stringify(
            data,
            null,
            2
          )}`,
        },
      ],
    };
  }

  /*****************************************************************
   *  The server getter - defines all your tools (same as before).  *
   *****************************************************************/
  get server() {
    const server = new McpServer({
      name: "Microsoft Graph Service",
      version: "1.0.0",
    });

    /* -------------------------- YOUR TOOLS -------------------------- */

    server.tool(
      "microsoft-graph-api",
      "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management.",
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
          const data = await this.graphService.genericGraphRequest(
            params.path,
            params.method,
            params.body,
            params.queryParams,
            params.graphApiVersion,
            params.fetchAll,
            params.consistencyLevel
          );

          return this.formatResponse(
            `${params.method.toUpperCase()} ${params.path}`,
            data
          );
        } catch (err: unknown) {
          logger.error("Error in microsoft-graph-api tool", err);
          throw err;
        }
      }
    );

    /* ---- Example of your other specific tools, unchanged ---- */
    server.tool(
      "getCurrentUserProfile",
      "Get the current user's Microsoft Graph profile",
      {},
      async () => {
        const profile = await this.graphService.getCurrentUserProfile();
        return this.formatResponse("User profile retrieved", profile);
      }
    );

    server.tool(
      "getUsers",
      "Get users from Microsoft Graph",
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) => {
        const users = await this.graphService.getUsers(queryParams, fetchAll);
        return this.formatResponse("Users retrieved", users);
      }
    );

    server.tool(
      "getGroups",
      "Get groups from Microsoft Graph",
      {
        queryParams: z.record(z.string()).optional(),
        fetchAll: z.boolean().optional().default(false),
      },
      async ({ queryParams, fetchAll }) => {
        const groups = await this.graphService.getGroups(queryParams, fetchAll);
        return this.formatResponse("Groups retrieved", groups);
      }
    );

    /* … add the rest of your tools here (unchanged) … */

    return server;
  }
}
