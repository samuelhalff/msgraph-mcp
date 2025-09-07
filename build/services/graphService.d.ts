import { GraphUser, GraphMessage, GraphEvent } from '../types.js';
export declare class GraphService {
    private client;
    constructor(accessToken: string);
    getUserProfile(): Promise<GraphUser>;
    getMessages(top?: number): Promise<GraphMessage[]>;
    getEvents(top?: number): Promise<GraphEvent[]>;
    sendMail(subject: string, content: string, toRecipients: string[]): Promise<void>;
    createCalendarEvent(subject: string, start: string, end: string, attendees?: string[]): Promise<GraphEvent>;
}
