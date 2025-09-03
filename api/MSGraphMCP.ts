/* -------------------------------------------------------------------- *
 *  src/MSGraphMCP.ts – Lean, stateless per-request MCP HTTP handler     *
 * -------------------------------------------------------------------- */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { MSGraphService } from "./MSGraphService.js";
import { Env, MSGraphAuthContext } from "../types.js";
import logger from "./lib/logger.js";

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
      } as any);
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

  /** Build and configure the McpServer with all tools */
  get server() {
    const server = new McpServer({
      name: "Microsoft Graph Service",
      version: "1.0.0",
    });

    // 1) Universal Graph/Azure request
    server.registerTool(
      "microsoft-graph-api",
      {
        title: "Microsoft Graph API",
        description: "Versatile Graph / ARM request helper.",
        inputSchema: {
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
              "Graph API ConsistencyLevel header. Advised to be set to 'eventual' for Graph GET requests using advanced OData queries."
            ),
        },
      },
      async (p) => {
        try {
          if (p.apiType === "azure") {
            const res = await this.svc.azureRequest(
              p.path,
              p.method,
              p.body,
              p.queryParams,
              p.apiVersion,
              p.subscriptionId,
              p.fetchAll
            );
            return this.formatResponse(
              `Azure ${p.method.toUpperCase()} ${p.path}`,
              res
            );
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
          return this.formatResponse(
            `Graph ${p.method.toUpperCase()} ${p.path}`,
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
    server.registerTool(
      "getCurrentUserProfile",
      {
        title: "Get Current User Profile",
        description: "Get the current user's Microsoft Graph profile",
        inputSchema: {},
      },
      async () => {
        const res = await this.svc.getCurrentUserProfile();
        return this.formatResponse("User profile retrieved", res);
      }
    );

    server.registerTool(
      "getUsers",
      {
        title: "Get Users",
        description: "Get users from Microsoft Graph",
        inputSchema: {
          queryParams: z.record(z.string()).optional(),
          fetchAll: z.boolean().optional().default(false),
        },
      },
      async ({ queryParams, fetchAll }) => {
        const res = await this.svc.getUsers(queryParams, fetchAll);
        return this.formatResponse("Users retrieved", res);
      }
    );

    server.registerTool(
      "getGroups",
      {
        title: "Get Groups",
        description: "Get groups from Microsoft Graph",
        inputSchema: {
          queryParams: z.record(z.string()).optional(),
          fetchAll: z.boolean().optional().default(false),
        },
      },
      async ({ queryParams, fetchAll }) => {
        const res = await this.svc.getGroups(queryParams, fetchAll);
        return this.formatResponse("Groups retrieved", res);
      }
    );

    server.registerTool(
      "getApplications",
      {
        title: "Get Applications",
        description: "Get applications from Microsoft Graph",
        inputSchema: {
          queryParams: z.record(z.string()).optional(),
          fetchAll: z.boolean().optional().default(false),
        },
      },
      async ({ queryParams, fetchAll }) => {
        const res = await this.svc.getApplications(queryParams, fetchAll);
        return this.formatResponse("Applications retrieved", res);
      }
    );

    // 3) Outlook draft
    server.registerTool(
      "createDraftEmail",
      {
        title: "Create Draft Email",
        description: "Create an Outlook draft (saved to Drafts)",
        inputSchema: {
          subject: z.string().optional(),
          body: z.string().optional(),
          contentType: z.enum(["Text", "HTML"]).optional().default("Text"),
          toRecipients: z.array(z.string()).optional(),
          ccRecipients: z.array(z.string()).optional(),
          bccRecipients: z.array(z.string()).optional(),
        },
      },
      async ({
        subject,
        body,
        contentType,
        toRecipients,
        ccRecipients,
        bccRecipients,
      }) => {
        try {
          const msg = {
            subject: subject ?? "",
            body: { contentType: contentType ?? "Text", content: body ?? "" },
            toRecipients: (toRecipients ?? []).map((a) => ({
              emailAddress: { address: a },
            })),
            ccRecipients: (ccRecipients ?? []).map((a) => ({
              emailAddress: { address: a },
            })),
            bccRecipients: (bccRecipients ?? []).map((a) => ({
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
          logger.error("createDraftEmail error", { msg, subject, contentType });
          throw new Error(msg);
        }
      }
    );

    // 4) Calendar helpers
    server.registerTool(
      "getUpcomingEvents",
      {
        title: "Get Upcoming Events",
        description:
          "Get upcoming calendar events for the current user from Microsoft Graph",
        inputSchema: {
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
        },
      },
      async ({ numberOfEvents, startDateTime }) => {
        try {
          const queryParams = {
            $top: String(numberOfEvents),
            $orderby: "start/dateTime",
            $filter: `start/dateTime ge '${
              startDateTime || new Date().toISOString()
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
            numberOfEvents,
            startDateTime,
          });
          throw new Error(msg);
        }
      }
    );

    server.registerTool(
      "createCalendarEvent",
      {
        title: "Create Calendar Event",
        description:
          "Create a calendar event (meeting) with other people in Microsoft Graph",
        inputSchema: {
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
        },
      },
      async (p) => {
        try {
          const event = {
            subject: p.subject,
            start: { dateTime: p.startTime, timeZone: "UTC" },
            end: { dateTime: p.endTime, timeZone: "UTC" },
            body: { contentType: "Text", content: p.body || "" },
            location: { displayName: p.location || "" },
            attendees: p.attendees.map((email) => ({
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
    server.registerTool(
      "search-files",
      {
        title: "Search Files",
        description: "Search for files across OneDrive, SharePoint, and Teams using Microsoft Graph Search API.",
        inputSchema: {
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
        }
      },
      async (params) => {
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
                    queryString: `${params.query} AND (${params.fileTypes.map(type => `filetype:${type}`).join(" OR ")})`
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
            results: searchResults.value?.[0]?.hitsContainers?.[0]?.hits?.map((hit: any) => ({
              name: hit.resource?.name || "Unknown",
              webUrl: hit.resource?.webUrl || "",
              lastModified: hit.resource?.lastModifiedDateTime || "",
              size: hit.resource?.size || 0,
              fileType: hit.resource?.file?.mimeType || "",
              summary: hit.summary || "",
              path: hit.resource?.parentReference?.path || "",
              createdBy: hit.resource?.createdBy?.user?.displayName || "",
              modifiedBy: hit.resource?.lastModifiedBy?.user?.displayName || "",
              downloadUrl: hit.resource?.["@microsoft.graph.downloadUrl"] || "",
              score: hit.score || 0
            })) || []
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
    server.registerTool(
      "get-schedule",
      {
        title: "Get Schedule",
        description: "Get the free/busy availability information for a collection of users, distributions lists, or resources for a specified time period.",
        inputSchema: {
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
        }
      },
      async (params) => {
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
            schedules: scheduleData.value?.map((schedule: any, index: number) => ({
              email: params.schedules[index],
              availabilityView: schedule.availabilityView || [],
              busyTimes: schedule.busyTimes?.map((busyTime: any) => ({
                start: busyTime.start?.dateTime || "",
                end: busyTime.end?.dateTime || "",
                status: busyTime.status || "busy"
              })) || [],
              workingHours: schedule.workingHours ? {
                daysOfWeek: schedule.workingHours.daysOfWeek || [],
                startTime: schedule.workingHours.startTime || "",
                endTime: schedule.workingHours.endTime || "",
                timeZone: schedule.workingHours.timeZone?.name || "UTC"
              } : null,
              freeBusyStatus: this.getFreeBusyInterpretation(schedule.availabilityView || [])
            })) || []
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

    logger.info("McpServer configured");
    return server;
  }
}
