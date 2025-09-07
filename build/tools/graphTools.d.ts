import { GraphService } from '../services/graphService.js';
export declare class GraphTools {
    getToolDefinitions(): ({
        name: string;
        description: string;
        inputSchema: {
            type: string;
            properties: {
                top?: undefined;
                subject?: undefined;
                content?: undefined;
                to?: undefined;
                start?: undefined;
                end?: undefined;
                attendees?: undefined;
            };
            required: never[];
        };
    } | {
        name: string;
        description: string;
        inputSchema: {
            type: string;
            properties: {
                top: {
                    type: string;
                    description: string;
                    default: number;
                };
                subject?: undefined;
                content?: undefined;
                to?: undefined;
                start?: undefined;
                end?: undefined;
                attendees?: undefined;
            };
            required: never[];
        };
    } | {
        name: string;
        description: string;
        inputSchema: {
            type: string;
            properties: {
                subject: {
                    type: string;
                    description: string;
                };
                content: {
                    type: string;
                    description: string;
                };
                to: {
                    type: string;
                    items: {
                        type: string;
                    };
                    description: string;
                };
                top?: undefined;
                start?: undefined;
                end?: undefined;
                attendees?: undefined;
            };
            required: string[];
        };
    } | {
        name: string;
        description: string;
        inputSchema: {
            type: string;
            properties: {
                subject: {
                    type: string;
                    description: string;
                };
                start: {
                    type: string;
                    description: string;
                };
                end: {
                    type: string;
                    description: string;
                };
                attendees: {
                    type: string;
                    items: {
                        type: string;
                    };
                    description: string;
                };
                top?: undefined;
                content?: undefined;
                to?: undefined;
            };
            required: string[];
        };
    })[];
    executeTool(toolName: string, args: Record<string, unknown>, graphService: GraphService): Promise<unknown>;
}
