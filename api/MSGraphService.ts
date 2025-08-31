/* eslint-disable @typescexport interface MSGraphAuthContext {
  accessToken: string
  refreshToken?: string
  expiresIn?: number
  tokenType?: string
  scope?: string
}

export class AuthManager {
  private config: AuthConfig;
  private credential: any;
  private accessToken: string | null = null;
  private refreshToken: string | null = null;
  private tokenExpiration: Date | null = null;

  constructor(config: AuthConfig) {
    this.config = config;
  }-explicit-any */
import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import { InteractiveBrowserCredential, ClientSecretCredential, ClientCertificateCredential } from "@azure/identity";
import logger from "./lib/logger.js";
import { Env } from "../types";

// Note: In Cloudflare Workers, fetch is already available globally
// No need to set up isomorphic-fetch

export enum AuthMode {
  ClientCredentials = "ClientCredentials",
  Certificate = "Certificate",
  Interactive = "Interactive",
  ClientProvidedToken = "ClientProvidedToken"
}

export interface AuthConfig {
  mode: AuthMode;
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  accessToken?: string;
  redirectUri?: string;
  certificatePath?: string;
  certificatePassword?: string;
}

export interface MSGraphAuthContext {
  accessToken: string;
  refreshToken?: string;
}

export class AuthManager {
  private config: AuthConfig;
  private credential: any;
  private accessToken: string | null = null;
  private refreshToken: string | null = null;
  private tokenExpiration: Date | null = null;

  constructor(config: AuthConfig) {
    this.config = config;
  }

  async initialize() {
    switch (this.config.mode) {
      case AuthMode.ClientCredentials:
        this.credential = new ClientSecretCredential(
          this.config.tenantId!,
          this.config.clientId!,
          this.config.clientSecret!
        );
        break;
      case AuthMode.Certificate:
        if (this.config.certificatePath) {
          this.credential = new ClientCertificateCredential(
            this.config.tenantId!,
            this.config.clientId!,
            this.config.certificatePath
          );
        } else {
          throw new Error("Certificate path is required for certificate authentication");
        }
        break;
      case AuthMode.Interactive:
        // For Cloudflare Workers, interactive auth might need different handling
        // This is a simplified version
        this.credential = new InteractiveBrowserCredential({
          tenantId: this.config.tenantId,
          clientId: this.config.clientId,
          redirectUri: this.config.redirectUri
        });
        break;
      case AuthMode.ClientProvidedToken:
        if (this.config.accessToken) {
          this.accessToken = this.config.accessToken;
        }
        break;
    }
  }

  getAuthMode(): AuthMode {
    return this.config.mode;
  }

  getAzureCredential() {
    return this.credential;
  }

  getGraphAuthProvider() {
    return {
      getAccessToken: async () => {
        if (this.config.mode === AuthMode.ClientProvidedToken) {
          if (!this.accessToken) {
            throw new Error("No access token available");
          }
          return this.accessToken;
        }

        const tokenResponse = await this.credential.getToken("https://graph.microsoft.com/.default");
        if (!tokenResponse || !tokenResponse.token) {
          throw new Error("Failed to acquire access token");
        }
        return tokenResponse.token;
      }
    };
  }

  updateAccessToken(token: string, expiration?: Date, refreshToken?: string) {
    this.accessToken = token;
    this.tokenExpiration = expiration || new Date(Date.now() + 3600000); // 1 hour default
    if (refreshToken) {
      this.refreshToken = refreshToken;
    }
  }

  async refreshAccessToken(): Promise<void> {
    if (!this.refreshToken) {
      throw new Error('No refresh token available');
    }

    const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: this.refreshToken,
        client_id: this.config.clientId!,
        client_secret: this.config.clientSecret || '',
        scope: 'https://graph.microsoft.com/.default'
      })
    });

    if (!response.ok) {
      throw new Error(`Failed to refresh access token: ${response.status} ${response.statusText}`);
    }

    const data = await response.json() as {
      access_token: string;
      refresh_token?: string;
      expires_in: number;
      scope: string;
      token_type: string;
    };
    this.accessToken = data.access_token;
    if (data.refresh_token) {
      this.refreshToken = data.refresh_token;
    }
    this.tokenExpiration = new Date(Date.now() + (data.expires_in * 1000));
  }

  async getTokenStatus() {
    if (this.config.mode === AuthMode.ClientProvidedToken) {
      return {
        isExpired: this.tokenExpiration ? new Date() > this.tokenExpiration : false,
        scopes: ["User provided token - scopes unknown"],
        expiresOn: this.tokenExpiration?.toISOString()
      };
    }

    try {
      const token = await this.credential.getToken("https://graph.microsoft.com/.default");
      return {
        isExpired: token.expiresOnTimestamp ? Date.now() > token.expiresOnTimestamp * 1000 : false,
        scopes: ["https://graph.microsoft.com/.default"],
        expiresOn: token.expiresOnTimestamp ? new Date(token.expiresOnTimestamp * 1000).toISOString() : undefined
      };
    } catch (error) {
      return {
        isExpired: true,
        error: error instanceof Error ? error.message : String(error)
      };
    }
  }
}

export class MSGraphService {
    private env: Env
    private authManager: AuthManager
    private graphClient: Client | null = null
    private useGraphBeta: boolean

    constructor(env: Env, authContext: MSGraphAuthContext, authConfig: AuthConfig) {
        this.env = env
        this.useGraphBeta = (this.env as any).USE_GRAPH_BETA !== 'false'
        this.authManager = new AuthManager(authConfig)
        
        if (authConfig.mode === AuthMode.ClientProvidedToken && authContext.accessToken) {
            this.authManager.updateAccessToken(authContext.accessToken, undefined, authContext.refreshToken)
        }
    }

    async initialize() {
        await this.authManager.initialize()
        const authProvider = this.authManager.getGraphAuthProvider()
        this.graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
        })
    }

    private getGraphClient(): Client {
        if (!this.graphClient) {
            throw new Error("Graph client not initialized")
        }
        return this.graphClient
    }

    async makeGraphRequest(
        path: string, 
        method: 'get' | 'post' | 'put' | 'patch' | 'delete' = 'get',
        body?: any,
        queryParams?: Record<string, string>,
        graphApiVersion?: 'v1.0' | 'beta',
        fetchAll: boolean = false,
        consistencyLevel?: string
    ): Promise<any> {
        const effectiveVersion = graphApiVersion || (this.useGraphBeta ? 'beta' : 'v1.0')
        
        try {
            let request = this.getGraphClient().api(path).version(effectiveVersion)

            if (queryParams && Object.keys(queryParams).length > 0) {
                request = request.query(queryParams)
            }

            if (consistencyLevel) {
                request = request.header('ConsistencyLevel', consistencyLevel)
            }

            switch (method.toLowerCase()) {
                case 'get':
                    if (fetchAll) {
                        const firstPageResponse: PageCollection = await request.get()
                        const odataContext = firstPageResponse['@odata.context']
                        const allItems: any[] = firstPageResponse.value || []
                        
                        const callback = (item: any) => {
                            allItems.push(item)
                            return true
                        }

                        const pageIterator = new PageIterator(this.getGraphClient(), firstPageResponse, callback)
                        await pageIterator.iterate()

                        return {
                            '@odata.context': odataContext,
                            value: allItems
                        }
                    } else {
                        return await request.get()
                    }
                case 'post':
                    return await request.post(body ?? {})
                case 'put':
                    return await request.put(body ?? {})
                case 'patch':
                    return await request.patch(body ?? {})
                case 'delete':
                    const result = await request.delete()
                    return result === undefined || result === null ? { status: "Success (No Content)" } : result
                default:
                    throw new Error(`Unsupported method: ${method}`)
            }
        } catch (error: any) {
            // Handle 401 Unauthorized - try to refresh token and retry
            if (error.statusCode === 401 || error.code === 'InvalidAuthenticationToken' || 
                (error.message && error.message.includes('401'))) {
                
                logger.info('Received 401 error, attempting to refresh token...')
                
                try {
                    // Try to refresh the token
                    await this.authManager.refreshAccessToken()
                    
                    // Reinitialize the Graph client with new token
                    const authProvider = this.authManager.getGraphAuthProvider()
                    this.graphClient = Client.initWithMiddleware({
                        authProvider: authProvider,
                    })
                    
                    // Retry the request with refreshed token
                    let request = this.getGraphClient().api(path).version(effectiveVersion)

                    if (queryParams && Object.keys(queryParams).length > 0) {
                        request = request.query(queryParams)
                    }

                    if (consistencyLevel) {
                        request = request.header('ConsistencyLevel', consistencyLevel)
                    }

                    switch (method.toLowerCase()) {
                        case 'get':
                            if (fetchAll) {
                                const firstPageResponse: PageCollection = await request.get()
                                const odataContext = firstPageResponse['@odata.context']
                                const allItems: any[] = firstPageResponse.value || []
                                
                                const callback = (item: any) => {
                                    allItems.push(item)
                                    return true
                                }

                                const pageIterator = new PageIterator(this.getGraphClient(), firstPageResponse, callback)
                                await pageIterator.iterate()

                                return {
                                    '@odata.context': odataContext,
                                    value: allItems
                                }
                            } else {
                                return await request.get()
                            }
                        case 'post':
                            return await request.post(body ?? {})
                        case 'put':
                            return await request.put(body ?? {})
                        case 'patch':
                            return await request.patch(body ?? {})
                        case 'delete':
                            const result = await request.delete()
                            return result === undefined || result === null ? { status: "Success (No Content)" } : result
                        default:
                            throw new Error(`Unsupported method: ${method}`)
                    }
                } catch (refreshError) {
                    logger.error('Failed to refresh token:', refreshError)
                    throw new Error(`Authentication failed: ${refreshError instanceof Error ? refreshError.message : String(refreshError)}`)
                }
            }
            
            // Re-throw the original error if it's not a 401
            throw error
        }
    }

    // User methods
    async getCurrentUserProfile(): Promise<any> {
        return this.makeGraphRequest('/me')
    }

    async getUsers(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/users', 'get', undefined, queryParams, undefined, fetchAll)
    }

    async getUser(userId: string): Promise<any> {
        return this.makeGraphRequest(`/users/${userId}`)
    }

    // Groups methods
    async getGroups(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/groups', 'get', undefined, queryParams, undefined, fetchAll)
    }

    async getGroup(groupId: string): Promise<any> {
        return this.makeGraphRequest(`/groups/${groupId}`)
    }

    async getGroupMembers(groupId: string, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest(`/groups/${groupId}/members`, 'get', undefined, undefined, undefined, fetchAll)
    }

    // Directory objects
    async getDirectoryObjects(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/directoryObjects', 'get', undefined, queryParams, undefined, fetchAll)
    }

    // Applications
    async getApplications(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/applications', 'get', undefined, queryParams, undefined, fetchAll)
    }

    async getApplication(appId: string): Promise<any> {
        return this.makeGraphRequest(`/applications/${appId}`)
    }

    // Service principals
    async getServicePrincipals(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/servicePrincipals', 'get', undefined, queryParams, undefined, fetchAll)
    }

    // Mail methods (if needed)
    async getMessages(userId: string = 'me', queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest(`/users/${userId}/messages`, 'get', undefined, queryParams, undefined, fetchAll)
    }

    async sendMail(userId: string = 'me', message: any): Promise<any> {
        return this.makeGraphRequest(`/users/${userId}/sendMail`, 'post', { message })
    }

    // Calendar methods
    async getEvents(userId: string = 'me', queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest(`/users/${userId}/events`, 'get', undefined, queryParams, undefined, fetchAll)
    }

    async createEvent(userId: string = 'me', event: any): Promise<any> {
        return this.makeGraphRequest(`/users/${userId}/events`, 'post', event)
    }

    // OneDrive/SharePoint methods
    async getDriveItems(userId: string = 'me', itemId?: string, queryParams?: Record<string, string>): Promise<any> {
        const path = itemId ? `/users/${userId}/drive/items/${itemId}` : `/users/${userId}/drive/root/children`
        return this.makeGraphRequest(path, 'get', undefined, queryParams)
    }

    // Teams methods
    async getTeams(queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest('/teams', 'get', undefined, queryParams, undefined, fetchAll)
    }

    async getTeamChannels(teamId: string, queryParams?: Record<string, string>, fetchAll: boolean = false): Promise<any> {
        return this.makeGraphRequest(`/teams/${teamId}/channels`, 'get', undefined, queryParams, undefined, fetchAll)
    }

    // Generic method for any Graph API endpoint
    async genericGraphRequest(
        path: string,
        method: 'get' | 'post' | 'put' | 'patch' | 'delete' = 'get',
        body?: any,
        queryParams?: Record<string, string>,
        graphApiVersion?: 'v1.0' | 'beta',
        fetchAll: boolean = false,
        consistencyLevel?: string
    ): Promise<any> {
        return this.makeGraphRequest(path, method, body, queryParams, graphApiVersion, fetchAll, consistencyLevel)
    }
} 