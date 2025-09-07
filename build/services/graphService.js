import axios from 'axios';
import { logger } from '../utils/logger.js';
const log = logger('graphService');
export class GraphService {
    client;
    constructor(accessToken) {
        this.client = axios.create({
            baseURL: 'https://graph.microsoft.com/v1.0',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
    }
    async getUserProfile() {
        log.info('Fetching user profile');
        const response = await this.client.get('/me');
        return response.data;
    }
    async getMessages(top = 10) {
        log.info(`Fetching top ${top} messages`);
        const response = await this.client.get(`/me/messages?$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`);
        return response.data.value;
    }
    async getEvents(top = 10) {
        log.info(`Fetching top ${top} calendar events`);
        const response = await this.client.get(`/me/events?$top=${top}&$select=id,subject,start,end,organizer`);
        return response.data.value;
    }
    async sendMail(subject, content, toRecipients) {
        log.info(`Sending email with subject: ${subject}`);
        const message = {
            subject,
            body: {
                contentType: 'Text',
                content
            },
            toRecipients: toRecipients.map(email => ({
                emailAddress: { address: email }
            }))
        };
        await this.client.post('/me/sendMail', { message });
    }
    async createCalendarEvent(subject, start, end, attendees) {
        log.info(`Creating calendar event: ${subject}`);
        const event = {
            subject,
            start: {
                dateTime: start,
                timeZone: 'UTC'
            },
            end: {
                dateTime: end,
                timeZone: 'UTC'
            },
            attendees: attendees?.map(email => ({
                emailAddress: { address: email },
                type: 'required'
            })) || []
        };
        const response = await this.client.post('/me/events', event);
        return response.data;
    }
}
