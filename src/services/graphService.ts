import axios from 'axios';
import type { GraphUser, GraphMessage, GraphEvent } from '../types.d.ts';
import { logger } from '../utils/logger.ts';

const log = logger('graphService');

export class GraphService {
  private client: ReturnType<typeof axios.create>;

  constructor(accessToken: string) {
    this.client = axios.create({
      baseURL: 'https://graph.microsoft.com/v1.0',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });
  }

  async getUserProfile(): Promise<GraphUser> {
    log.info('Fetching user profile');
    const response = await this.client.get<GraphUser>('/me');
    return response.data;
  }

  async getMessages(top: number = 10): Promise<GraphMessage[]> {
    log.info(`Fetching top ${top} messages`);
    const response = await this.client.get<{ value: GraphMessage[] }>(`/me/messages?$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`);
    return response.data.value;
  }

  async getEvents(top: number = 10): Promise<GraphEvent[]> {
    log.info(`Fetching top ${top} calendar events`);
    const response = await this.client.get<{ value: GraphEvent[] }>(`/me/events?$top=${top}&$select=id,subject,start,end,organizer`);
    return response.data.value;
  }

  async sendMail(subject: string, content: string, toRecipients: string[]): Promise<void> {
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

  async createCalendarEvent(subject: string, start: string, end: string, attendees?: string[]): Promise<GraphEvent> {
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

    const response = await this.client.post<GraphEvent>('/me/events', event);
    return response.data;
  }
}