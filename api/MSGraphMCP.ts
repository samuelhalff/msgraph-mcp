import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { MSGraphService } from "./MSGraphService.js";
import { Env, MSGraphAuthContext } from "../types.js";
import logger from "./lib/logger.js";
import { throttlingManager } from "./lib/throttling-manager.js";
import { z, ZodObject, ZodTypeAny } from "zod";

// Tool parameter interfaces
interface GraphApiParams {
  apiType?: "graph" | "azure";
  path: string;
  method?: string;
  body?: Record<string, unknown>;
  queryParams?: Record<string, string>;
  apiVersion?: string;
  subscriptionId?: string;
  graphApiVersion?: "v1.0" | "beta";
  fetchAll?: boolean;
  consistencyLevel?: string;
}

interface UserGroupParams {
  queryParams?: Record<string, string>;
  fetchAll?: boolean;
}

interface SearchUsersParams {
  query: string;
}

interface SendMailParams {
  to: string;
  subject: string;
  body: string;
}

interface ListCalendarEventsParams {
  startDateTime?: string;
  endDateTime?: string;
}

interface CreateCalendarEventParams {
  subject: string;
  start: string;
  end: string;
  attendees?: string[];
  body?: string;
  location?: string;
}

interface DraftEmailParams {
  subject?: string;
  body?: string;
  contentType?: "Text" | "HTML";
  toRecipients?: string[];
  ccRecipients?: string[];
  bccRecipients?: string[];
}

interface UpcomingEventsParams {
  numberOfEvents?: number;
  startDateTime?: string;
}

interface CalendarEventParams {
  subject: string;
  startTime: string;
  endTime: string;
  attendees: string[];
  body?: string;
  location?: string;
  isOnlineMeeting?: boolean;
}

interface SearchFilesParams {
  query: string;
  entityTypes?: string[];
  size?: number;
  from?: number;
  fileTypes?: string[];
  contentSource?: "default" | "sharepoint" | "onedrive";
  sortBy?: "relevance" | "lastModifiedDateTime" | "name" | "size";
  sortOrder?: "asc" | "desc";
}

interface GetScheduleParams {
  schedules: string[];
  startTime: string;
  endTime: string;
  availabilityViewInterval?: number;
}

export class MSGraphMCP {
  public readonly env: Env;
  public readonly auth: MSGraphAuthContext;

  constructor(env: Env, auth: MSGraphAuthContext) {
    this.env = env;
    this.auth = auth;
  }

  
  #svc?: MSGraphService;
  private get svc(): MSGraphService {
    if (!this.#svc) {
      logger.info("Creating new MSGraphService instance");
      
      // Validate required environment variables
      if (!this.env.TENANT_ID || !this.env.CLIENT_ID) {
        throw new Error('TENANT_ID and CLIENT_ID are required');
      }
      
      this.#svc = new MSGraphService(this.env, this.auth, {
        tenantId: this.env.TENANT_ID,
        clientId: this.env.CLIENT_ID,
        clientSecret: this.env.CLIENT_SECRET,
        mode:
          this.env.USE_CLIENT_TOKEN === "true"
            ? "ClientProvidedToken"
            : this.env.USE_CERTIFICATE === "true"
            ? "Certificate"
            : this.env.USE_INTERACTIVE === "true"
            ? "Interactive"
            : "ClientCredentials",
        redirectUri: this.env.REDIRECT_URI,
        certificatePath: this.env.CERTIFICATE_PATH,
        certificatePassword: this.env.CERTIFICATE_PASSWORD,
      });
    }
    return this.#svc;
  }

  /** Helper to format tool results as MCP content */
  formatResponse(label: string, data: unknown) {
    const text =
      typeof data === "string" ? data : JSON.stringify(data, null, 2);
    return {
      content: [
        {
          type: "text" as const,
          text: `Success! ${label}\n\nResult:\n${text}`,
        },
      ],
    };
  }

  /** Helper to interpret availability view array into human-readable status */
  private getFreeBusyInterpretation(availabilityView: string[]): string {
    if (!availabilityView || availabilityView.length === 0) {
      return "No availability data";
    }

    const statusCounts = availabilityView.reduce((acc, status) => {
      acc[status] = (acc[status] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    const total = availabilityView.length;
    const statusPercentages = Object.entries(statusCounts).map(([status, count]) => ({
      status: this.getStatusLabel(status),
      percentage: Math.round((count / total) * 100)
    }));

    return statusPercentages
      .sort((a, b) => b.percentage - a.percentage)
      .map(({ status, percentage }) => `${status}: ${percentage}%`)
      .join(", ");
  }

  /** Convert numeric status to human-readable label */
  private getStatusLabel(status: string): string {
    switch (status) {
      case "0": return "Free";
      case "1": return "Tentative";
      case "2": return "Busy";
      case "3": return "Out of Office";
      case "4": return "Working Elsewhere";
      default: return `Unknown (${status})`;
    }
  }

  // Track tools as they're registered with the server
  private toolRegistry = new Map<string, { name: string; description: string; inputSchema: ZodTypeAny }>();
  private toolHandlers = new Map<string, (args: Record<string, unknown>) => Promise<{ content: Array<{ type: "text", text: string }> }> | { content: Array<{ type: "text", text: string }> }>();

  /** Register a tool with tracking */
  private registerServerTool(
    server: McpServer, 
    name: string, 
    schema: { title?: string; description?: string; inputSchema: ZodTypeAny }, 
    handler: (args: Record<string, unknown>) => Promise<{ content: Array<{ type: "text", text: string }> }> | { content: Array<{ type: "text", text: string }> }
  ) {
    // Track the tool with full schema information for HTTP transport
    this.toolRegistry.set(name, {
      name,
      description: schema.description || schema.title || `Microsoft Graph tool: ${name}`,
      inputSchema: schema.inputSchema
    });
    
    // Store the handler for direct tool calls
    this.toolHandlers.set(name, handler);
    
    // Register with the server using proper MCP format
    // MCP SDK expects the full schema object but with inputSchema as ZodRawShape
    const mcpSchema = {
      title: schema.title,
      description: schema.description,
      inputSchema: schema.inputSchema._def?.shape || (schema.inputSchema as ZodObject<Record<string, ZodTypeAny>>).shape || {}
    };
    server.registerTool(name, mcpSchema, handler);
  }

  /** Call a tool directly by name */
  async callTool(name: string, args: Record<string, unknown>) {
    const handler = this.toolHandlers.get(name);
    if (!handler) {
      throw new Error(`Tool '${name}' not found`);
    }
    return await handler(args);
  }

  /** Get list of available tools from our registry */
  getAvailableTools() {
    // If no tools registered yet, build the server to populate the registry
    if (this.toolRegistry.size === 0) {
      // Access server getter to trigger tool registration
      void this.server;
      // Server is now created and tools are registered
    }
    
    return Array.from(this.toolRegistry.values());
  }

  /** Get detailed tool information for debugging */
  getToolsDebugInfo() {
    // Ensure registry is populated
    if (this.toolRegistry.size === 0) {
      void this.server;
    }
    
    return {
      totalTools: this.toolRegistry.size,
      toolNames: Array.from(this.toolRegistry.keys()),
      tools: Array.from(this.toolRegistry.values()).map(tool => ({
        name: tool.name,
        description: tool.description,
        hasInputSchema: !!tool.inputSchema,
        inputSchemaType: tool.inputSchema?.constructor?.name || typeof tool.inputSchema,
        schemaKeys: tool.inputSchema && typeof tool.inputSchema === 'object' && 'shape' in tool.inputSchema 
          ? Object.keys((tool.inputSchema as ZodObject<Record<string, ZodTypeAny>>)._def.shape || {})
          : []
      })),
      hasHandlers: this.toolHandlers.size,
      handlerNames: Array.from(this.toolHandlers.keys())
    };
  }

  /** Check if a tool exists */
  hasTools(toolName: string): boolean {
    // Ensure registry is populated
    if (this.toolRegistry.size === 0) {
      // Access server getter to trigger tool registration
      void this.server;
      // Server is now created and tools are registered
    }
    
    return this.toolRegistry.has(toolName);
  }

  /** Build and configure the McpServer with all tools */
  get server() {
    const server = new McpServer({
      name: "Microsoft Graph Service",
      version: "1.0.0",
    });

    // 1) Universal Graph/Azure request
    this.registerServerTool(
      server,
      "microsoft-graph-api",
      {
        title: "Microsoft Graph API",
        description: "Versatile Graph / ARM request helper.",
        inputSchema: z.object({
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
            .optional()
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
              "Graph API ConsistencyLevel header. Advised to be set to 'eventual' for Graph GET requests using advanced OData queries."
            ),
        }),
      },
      async (args: Record<string, unknown>) => {
        const p = args as unknown as GraphApiParams;
        try {
          if (p.apiType === "azure") {
            const res = await this.svc.azureRequest(
              p.path,
              p.method || "get",
              p.body,
              p.queryParams,
              p.apiVersion,
              p.subscriptionId,
              p.fetchAll
            );
            return this.formatResponse(
              `Azure ${(p.method || "get").toUpperCase()} ${p.path}`,
              res
            );
          }

          const res = await this.svc.genericGraphRequest(
            p.path,
            p.method || "get",
            p.body,
            p.queryParams,
            p.graphApiVersion,
            p.fetchAll,
            p.consistencyLevel
          );
          return this.formatResponse(
            `Graph ${(p.method || "get").toUpperCase()} ${p.path}`,
            res
          );
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("microsoft-graph-api tool error", { msg, params: p });
          throw new Error(msg);
        }
      }
    );

    // 2) Convenience tools
    this.registerServerTool(
      server,
      "microsoft-graph-profile",
      {
        title: "Get Current User Profile",
        description: "Get the current user's Microsoft Graph profile",
        inputSchema: z.object({}),
      },
      async () => {
        const res = await this.svc.getCurrentUserProfile();
        return this.formatResponse("User profile retrieved", res);
      }
    );

    this.registerServerTool(
      server,
      "list-users",
      {
        title: "Get Users",
        description: "Get users from Microsoft Graph",
        inputSchema: z.object({
          queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
          fetchAll: z.boolean().optional().default(false).describe("Fetch all pages of results")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as UserGroupParams;
        const res = await this.svc.getUsers(params.queryParams, params.fetchAll);
        return this.formatResponse("Users retrieved", res);
      }
    );

    this.registerServerTool(
      server,
      "list-groups",
      {
        title: "Get Groups",
        description: "Get groups from Microsoft Graph",
        inputSchema: z.object({
          queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
          fetchAll: z.boolean().optional().default(false).describe("Fetch all pages of results")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as UserGroupParams;
        const res = await this.svc.getGroups(params.queryParams, params.fetchAll);
        return this.formatResponse("Groups retrieved", res);
      }
    );

    // Add missing tools that are in callTool switch
    this.registerServerTool(
      server,
      "search-users",
      {
        title: "Search Users",
        description: "Search for users in Microsoft Graph",
        inputSchema: z.object({
          query: z.string().describe("Search query for users")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as SearchUsersParams;
        const res = await this.svc.getUsers({ $search: `"displayName:${params.query}"` });
        return this.formatResponse("User search completed", res);
      }
    );

    this.registerServerTool(
      server,
      "send-mail",
      {
        title: "Send Mail",
        description: "Send an email via Microsoft Graph",
        inputSchema: z.object({
          to: z.string().describe("Recipient email address"),
          subject: z.string().describe("Email subject"),
          body: z.string().describe("Email body")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as SendMailParams;
        const res = await this.svc.genericGraphRequest('/me/sendMail', 'post', {
          message: {
            subject: params.subject,
            body: { contentType: "Text", content: params.body },
            toRecipients: [{ emailAddress: { address: params.to } }]
          }
        });
        return this.formatResponse("Email sent", res);
      }
    );

    this.registerServerTool(
      server,
      "list-calendar-events",
      {
        title: "List Calendar Events",
        description: "List calendar events for the current user",
        inputSchema: z.object({
          startDateTime: z.string().optional().describe("Start date-time filter"),
          endDateTime: z.string().optional().describe("End date-time filter")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as ListCalendarEventsParams;
        const queryParams: Record<string, string> = {};
        if (params.startDateTime) queryParams.$filter = `start/dateTime ge '${params.startDateTime}'`;
        if (params.endDateTime) {
          const filter = queryParams.$filter ? `${queryParams.$filter} and end/dateTime le '${params.endDateTime}'` : `end/dateTime le '${params.endDateTime}'`;
          queryParams.$filter = filter;
        }
        queryParams.$orderby = "start/dateTime";
        queryParams.$top = "50";
        
        const res = await this.svc.genericGraphRequest('/me/events', 'get', undefined, queryParams);
        return this.formatResponse("Upcoming events retrieved", res);
      }
    );

    this.registerServerTool(
      server,
      "create-calendar-event",
      {
        title: "Create Calendar Event",
        description: "Create a new calendar event",
        inputSchema: z.object({
          subject: z.string().describe("Event subject"),
          start: z.string().describe("Start time in ISO format"),
          end: z.string().describe("End time in ISO format"),
          attendees: z.array(z.string()).optional().describe("Attendee email addresses"),
          body: z.string().optional().describe("Event body"),
          location: z.string().optional().describe("Event location")
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as CreateCalendarEventParams;
        const event: Record<string, unknown> = {
          subject: params.subject,
          start: { dateTime: params.start, timeZone: "UTC" },
          end: { dateTime: params.end, timeZone: "UTC" }
        };
        
        if (params.body) event.body = { contentType: "HTML", content: params.body };
        if (params.location) event.location = { displayName: params.location };
        if (params.attendees) {
          event.attendees = params.attendees.map((email: string) => ({
            emailAddress: { address: email },
            type: "required"
          }));
        }
        
        const res = await this.svc.genericGraphRequest('/me/events', 'post', event);
        return this.formatResponse("Calendar event created", res);
      }
    );

    this.registerServerTool(
      server,
      "getApplications",
      {
        title: "Get Applications",
        description: "Get applications from Microsoft Graph",
        inputSchema: z.object({
          queryParams: z.record(z.string()).optional(),
          fetchAll: z.boolean().optional().default(false),
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as UserGroupParams;
        const res = await this.svc.getApplications(params.queryParams, params.fetchAll);
        return this.formatResponse("Applications retrieved", res);
      }
    );

    // 3) Outlook draft
    this.registerServerTool(
      server,
      "createDraftEmail",
      {
        title: "Create Draft Email",
        description: "Create an Outlook draft (saved to Drafts)",
        inputSchema: z.object({
          subject: z.string().optional(),
          body: z.string().optional(),
          contentType: z.enum(["Text", "HTML"]).optional().default("Text"),
          toRecipients: z.array(z.string()).optional(),
          ccRecipients: z.array(z.string()).optional(),
          bccRecipients: z.array(z.string()).optional(),
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as DraftEmailParams;
        try {
          const msg = {
            subject: params.subject ?? "",
            body: { contentType: params.contentType ?? "Text", content: params.body ?? "" },
            toRecipients: (params.toRecipients ?? []).map((a: string) => ({
              emailAddress: { address: a },
            })),
            ccRecipients: (params.ccRecipients ?? []).map((a: string) => ({
              emailAddress: { address: a },
            })),
            bccRecipients: (params.bccRecipients ?? []).map((a: string) => ({
              emailAddress: { address: a },
            })),
          };
          const res = await this.svc.genericGraphRequest(
            "/me/messages",
            "post",
            msg
          );
          return this.formatResponse("Draft created", res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("createDraftEmail error", { msg, subject: params.subject, contentType: params.contentType });
          throw new Error(msg);
        }
      }
    );

    // 4) Calendar helpers
    this.registerServerTool(
      server,
      "getUpcomingEvents",
      {
        title: "Get Upcoming Events",
        description:
          "Get upcoming calendar events for the current user from Microsoft Graph",
        inputSchema: z.object({
          numberOfEvents: z
            .number()
            .optional()
            .default(10)
            .describe("Number of events to retrieve. Default: 10"),
          startDateTime: z
            .string()
            .optional()
            .describe(
              "Start date-time in ISO format to filter events from. Default: current time"
            ),
        }),
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as UpcomingEventsParams;
        try {
          const queryParams = {
            $top: String(params.numberOfEvents ?? 10),
            $orderby: "start/dateTime",
            $filter: `start/dateTime ge '${
              params.startDateTime || new Date().toISOString()
            }'`,
          } as Record<string, string>;
          const res = await this.svc.genericGraphRequest(
            "/me/events",
            "get",
            null,
            queryParams,
            "v1.0",
            true
          );
          return this.formatResponse("Upcoming events retrieved", res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("getUpcomingEvents tool error", {
            msg,
            numberOfEvents: params.numberOfEvents,
            startDateTime: params.startDateTime,
          });
          throw new Error(msg);
        }
      }
    );

    this.registerServerTool(
      server,
      "createCalendarEvent",
      {
        title: "Create Calendar Event",
        description:
          "Create a calendar event (meeting) with other people in Microsoft Graph",
        inputSchema: z.object({
          subject: z.string().describe("Subject of the event"),
          startTime: z
            .string()
            .describe(
              "Start time in ISO format (e.g., '2025-09-02T10:00:00Z')"
            ),
          endTime: z
            .string()
            .describe("End time in ISO format (e.g., '2025-09-02T11:00:00Z')"),
          attendees: z
            .array(z.string())
            .describe("Array of email addresses of attendees"),
          body: z.string().optional().describe("Body content of the event"),
          location: z.string().optional().describe("Location of the event"),
          isOnlineMeeting: z
            .boolean()
            .optional()
            .default(false)
            .describe("Whether it's an online meeting (e.g., Teams)"),
        }),
      },
      async (args: Record<string, unknown>) => {
        const p = args as unknown as CalendarEventParams;
        try {
          const event = {
            subject: p.subject,
            start: { dateTime: p.startTime, timeZone: "UTC" },
            end: { dateTime: p.endTime, timeZone: "UTC" },
            body: { contentType: "Text", content: p.body || "" },
            location: { displayName: p.location || "" },
            attendees: p.attendees.map((email: string) => ({
              emailAddress: { address: email },
              type: "required",
            })),
            isOnlineMeeting: p.isOnlineMeeting,
          };
          const res = await this.svc.genericGraphRequest(
            "/me/events",
            "post",
            event
          );
          return this.formatResponse("Calendar event created", res);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("createCalendarEvent tool error", { msg, params: p });
          throw new Error(msg);
        }
      }
    );

    // Search Files tool - Microsoft Graph Search API for files
    this.registerServerTool(
      server,
      "search-files",
      {
        title: "Search Files",
        description: "Search for files across OneDrive, SharePoint, and Teams using Microsoft Graph Search API.",
        inputSchema: z.object({
          query: z
            .string()
            .describe("Search query for files (e.g., 'quarterly report', 'meeting notes', or 'filename:document.pdf')"),
          entityTypes: z
            .array(z.enum(["driveItem"]))
            .optional()
            .default(["driveItem"])
            .describe("Entity types to search. For files, use 'driveItem'"),
          size: z
            .number()
            .min(1)
            .max(1000)
            .optional()
            .default(25)
            .describe("Number of results to return (1-1000, default: 25)"),
          from: z
            .number()
            .min(0)
            .optional()
            .default(0)
            .describe("Starting point for results (for pagination, default: 0)"),
          fileTypes: z
            .array(z.string())
            .optional()
            .describe("Filter by file types (e.g., ['pdf', 'docx', 'xlsx'])"),
          contentSource: z
            .enum(["default", "sharepoint", "onedrive"])
            .optional()
            .default("default")
            .describe("Content source to search: 'default' (all), 'sharepoint', or 'onedrive'"),
          sortBy: z
            .enum(["relevance", "lastModifiedDateTime", "name", "size"])
            .optional()
            .default("relevance")
            .describe("Sort results by: 'relevance', 'lastModifiedDateTime', 'name', or 'size'"),
          sortOrder: z
            .enum(["asc", "desc"])
            .optional()
            .default("desc")
            .describe("Sort order: 'asc' or 'desc'")
        })
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as SearchFilesParams;
        try {
          logger.info("Search files tool called", { params });

          // Build the search request according to Microsoft Graph Search API
          const searchRequest = {
            requests: [
              {
                entityTypes: params.entityTypes || ["driveItem"],
                query: {
                  queryString: params.query
                },
                from: params.from || 0,
                size: params.size || 25,
                sortProperties: [
                  {
                    name: params.sortBy || "relevance",
                    isDescending: (params.sortOrder || "desc") === "desc"
                  }
                ],
                ...(params.contentSource && params.contentSource !== "default" && {
                  contentSources: [params.contentSource]
                }),
                ...(params.fileTypes && params.fileTypes.length > 0 && {
                  query: {
                    queryString: `${params.query} AND (${params.fileTypes.map((type: string) => `filetype:${type}`).join(" OR ")})`
                  }
                })
              }
            ]
          };

          // Make the search request
          const searchResults = await this.svc.genericGraphRequest(
            "/search/query",
            "post",
            searchRequest
          );

          // Format the results for better readability
          const formattedResults = {
            totalResults: searchResults.value?.[0]?.hitsContainers?.[0]?.total || 0,
            results: searchResults.value?.[0]?.hitsContainers?.[0]?.hits?.map((hit: Record<string, unknown>) => {
              const resource = hit.resource as Record<string, unknown>;
              const file = resource?.file as Record<string, unknown>;
              const parentRef = resource?.parentReference as Record<string, unknown>;
              const createdBy = resource?.createdBy as Record<string, unknown>;
              const lastModifiedBy = resource?.lastModifiedBy as Record<string, unknown>;
              const createdByUser = createdBy?.user as Record<string, unknown>;
              const lastModifiedByUser = lastModifiedBy?.user as Record<string, unknown>;
              
              return {
                name: resource?.name as string || "Unknown",
                webUrl: resource?.webUrl as string || "",
                lastModified: resource?.lastModifiedDateTime as string || "",
                size: resource?.size as number || 0,
                fileType: file?.mimeType as string || "",
                summary: hit.summary as string || "",
                path: parentRef?.path as string || "",
                createdBy: createdByUser?.displayName as string || "",
                modifiedBy: lastModifiedByUser?.displayName as string || "",
                downloadUrl: (resource as Record<string, unknown>)?.["@microsoft.graph.downloadUrl"] as string || "",
                score: hit.score as number || 0
              };
            }) || []
          };

          return this.formatResponse(
            `Found ${formattedResults.totalResults} files matching "${params.query}"`,
            formattedResults
          );

        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("Search files tool error", { msg, params });
          throw new Error(`Failed to search files: ${msg}`);
        }
      }
    );

    // Get Schedule tool - Microsoft Graph Calendar getSchedule API
    this.registerServerTool(
      server,
      "get-schedule",
      {
        title: "Get Schedule",
        description: "Get the free/busy availability information for a collection of users, distributions lists, or resources for a specified time period.",
        inputSchema: z.object({
          schedules: z
            .array(z.string())
            .min(1)
            .max(20)
            .describe("Email addresses of users, distribution lists, or resources to get schedule for (max 20)"),
          startTime: z
            .string()
            .describe("Start time for the schedule query in ISO 8601 format (e.g., '2024-03-15T08:00:00.000Z')"),
          endTime: z
            .string()
            .describe("End time for the schedule query in ISO 8601 format (e.g., '2024-03-15T18:00:00.000Z')"),
          availabilityViewInterval: z
            .number()
            .min(5)
            .max(1440)
            .optional()
            .default(30)
            .describe("Interval in minutes for availability view (5-1440, default: 30). Represents the granularity of free/busy time.")
        })
      },
      async (args: Record<string, unknown>) => {
        const params = args as unknown as GetScheduleParams;
        try {
          logger.info("Get schedule tool called", { params });

          // Validate date format and order
          const startDate = new Date(params.startTime);
          const endDate = new Date(params.endTime);
          
          if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            throw new Error("Invalid date format. Use ISO 8601 format (e.g., '2024-03-15T08:00:00.000Z')");
          }
          
          if (startDate >= endDate) {
            throw new Error("Start time must be before end time");
          }

          // Check if time range is reasonable (not more than 62 days as per API limits)
          const daysDiff = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
          if (daysDiff > 62) {
            throw new Error("Time range cannot exceed 62 days");
          }

          // Build the request body according to Microsoft Graph API
          const requestBody = {
            schedules: params.schedules,
            startTime: {
              dateTime: params.startTime,
              timeZone: "UTC"
            },
            endTime: {
              dateTime: params.endTime,
              timeZone: "UTC"
            },
            availabilityViewInterval: params.availabilityViewInterval || 30
          };

          // Make the request to Microsoft Graph
          const scheduleData = await this.svc.genericGraphRequest(
            "/me/calendar/getSchedule",
            "post",
            requestBody
          );

          // Format the response for better readability
          const formattedSchedule = {
            queryPeriod: {
              startTime: params.startTime,
              endTime: params.endTime,
              intervalMinutes: params.availabilityViewInterval || 30
            },
            schedules: scheduleData.value?.map((schedule: Record<string, unknown>, index: number) => {
              const workingHours = schedule.workingHours as Record<string, unknown>;
              const timeZone = workingHours?.timeZone as Record<string, unknown>;
              
              return {
                email: params.schedules[index],
                availabilityView: schedule.availabilityView as string[] || [],
                busyTimes: (schedule.busyTimes as Record<string, unknown>[])?.map((busyTime: Record<string, unknown>) => {
                  const start = busyTime.start as Record<string, unknown>;
                  const end = busyTime.end as Record<string, unknown>;
                  return {
                    start: start?.dateTime as string || "",
                    end: end?.dateTime as string || "",
                    status: busyTime.status as string || "busy"
                  };
                }) || [],
                workingHours: workingHours ? {
                  daysOfWeek: workingHours.daysOfWeek as string[] || [],
                  startTime: workingHours.startTime as string || "",
                  endTime: workingHours.endTime as string || "",
                  timeZone: timeZone?.name as string || "UTC"
                } : null,
                freeBusyStatus: this.getFreeBusyInterpretation(schedule.availabilityView as string[] || [])
              };
            }) || []
          };

          return this.formatResponse(
            `Retrieved schedule information for ${params.schedules.length} recipient(s) from ${params.startTime} to ${params.endTime}`,
            formattedSchedule
          );

        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("Get schedule tool error", { msg, params });
          throw new Error(`Failed to get schedule: ${msg}`);
        }
      }
    );

    // Throttling Statistics tool - Monitor API throttling and performance
    this.registerServerTool(
      server,
      "throttling-stats",
      {
        title: "Throttling Statistics",
        description: "Get current throttling statistics and API performance metrics for Microsoft Graph requests.",
        inputSchema: z.object({})
      },
      async () => {
        try {
          logger.info("Throttling stats tool called");
          
          // Get current stats from throttling manager
          const stats = throttlingManager.getStats();
          
          const enhancedStats = {
            ...stats,
            timestamp: new Date().toISOString(),
            windowSize: "10 minutes",
            description: {
              totalRequests: "Total requests since server start",
              recentRequests: "Requests in the last 10 minutes",
              errorRate: "Percentage of failed requests (0.0 to 1.0)",
              throttledRequests: "Number of 429 (throttled) responses"
            }
          };

          return this.formatResponse("Throttling statistics retrieved", enhancedStats);

        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          logger.error("Throttling stats tool error", { msg });
          throw new Error(`Failed to get throttling stats: ${msg}`);
        }
      }
    );

    logger.info("McpServer configured");
    
    // Log all registered tools for debugging and verification
    const registeredTools = Array.from(this.toolRegistry.keys());
    logger.info("Registered MCP tools", { 
      count: registeredTools.length,
      tools: registeredTools,
      details: Array.from(this.toolRegistry.values()).map(tool => ({
        name: tool.name,
        description: tool.description,
        hasInputSchema: !!tool.inputSchema
      }))
    });
    
    return server;
  }
}
