export interface TokenData {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope: string;
}

export interface UserContext {
  userId: string;
  sessionId?: string;
}

export interface GraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
}

export interface GraphMessage {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
  receivedDateTime: string;
  bodyPreview: string;
}

export interface GraphEvent {
  id: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  organizer: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
}