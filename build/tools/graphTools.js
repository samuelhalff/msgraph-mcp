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
            },
            {
                name: 'create_calendar_event',
                description: 'Create a new calendar event',
                inputSchema: {
                    type: 'object',
                    properties: {
                        subject: {
                            type: 'string',
                            description: 'Event subject/title'
                        },
                        start: {
                            type: 'string',
                            description: 'Start date and time (ISO 8601 format)'
                        },
                        end: {
                            type: 'string',
                            description: 'End date and time (ISO 8601 format)'
                        },
                        attendees: {
                            type: 'array',
                            items: { type: 'string' },
                            description: 'Array of attendee email addresses (optional)'
                        }
                    },
                    required: ['subject', 'start', 'end']
                }
            }
        ];
    }
    async executeTool(toolName, args, graphService) {
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
            case 'create_calendar_event':
                const event = await graphService.createCalendarEvent(args.subject, args.start, args.end, args.attendees);
                return { success: true, event };
            default:
                throw new Error(`Unknown tool: ${toolName}`);
        }
    }
}
