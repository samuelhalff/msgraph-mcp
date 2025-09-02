/* eslint-disable @typescript-eslint/no-explicit-any */
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser/index.js';
import { ClientCertificateCredential, ClientSecretCredential } from '@azure/identity';
import { PublicClientApplication, InteractionType } from '@azure/msal-browser';
import { Env, MSGraphAuthContext } from '../types.js';
import logger from './lib/logger.js';

// Custom options interface to include all required properties
interface MSGraphServiceOptions extends ClientOptions, Partial<AuthCodeMSALBrowserAuthenticationProviderOptions> {
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  redirectUri?: string;
  certificatePath?: string;
  certificatePassword?: string;
  mode: 'ClientProvidedToken' | 'Certificate' | 'Interactive' | 'ClientCredentials';
}

export class MSGraphService {
  private client: Client;

  constructor(
    private env: Env,
    private auth: MSGraphAuthContext,
    options: MSGraphServiceOptions
  ) {
    logger.info('Initializing MSGraphService', { env: { ...env, CLIENT_SECRET: '[REDACTED]', ACCESS_TOKEN: '[REDACTED]' }, auth: { ...auth, accessToken: '[REDACTED]' } });

    if (!env.TENANT_ID || !env.CLIENT_ID) {
      logger.error('Missing required environment variables', { TENANT_ID: env.TENANT_ID, CLIENT_ID: env.CLIENT_ID });
      throw new Error('TENANT_ID and CLIENT_ID must be set');
    }

    let authProvider;

    if (options.mode === 'ClientProvidedToken') {
      logger.info('Using ClientProvidedToken mode');
      if (!auth.accessToken) {
        logger.error('No access token provided in ClientProvidedToken mode');
        throw new Error('Access token required for ClientProvidedToken mode');
      }
      authProvider = {
        getAccessToken: async () => {
          logger.info('Providing client-provided access token');
          return auth.accessToken;
        },
      };
    } else if (options.mode === 'Certificate') {
      logger.info('Using Certificate mode');
      if (!options.certificatePath || !options.certificatePassword) {
        logger.error('Certificate path or password missing');
        throw new Error('Certificate path and password required');
      }
      const credential = new ClientCertificateCredential(options.tenantId, options.clientId, {
        certificatePath: options.certificatePath,
        certificatePassword: options.certificatePassword,
      });
      authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: env.OAUTH_SCOPES?.split(' ') ?? ['https://graph.microsoft.com/.default'],
      });
    } else if (options.mode === 'Interactive') {
      logger.info('Using Interactive mode');
      // For browser-based authentication, we need to use MSAL browser library
      const msalConfig = {
        auth: {
          clientId: options.clientId,
          authority: `https://login.microsoftonline.com/${options.tenantId}`,
          redirectUri: options.redirectUri ?? env.REDIRECT_URI ?? 'http://mcp-server:3001/authorize',
        },
      };
      const msalInstance = new PublicClientApplication(msalConfig);
      authProvider = new AuthCodeMSALBrowserAuthenticationProvider(msalInstance, {
        scopes: env.OAUTH_SCOPES?.split(' ') ?? ['https://graph.microsoft.com/.default'],
        interactionType: InteractionType.Popup,
        account: null as any, // Will be set by MSAL during authentication
      });
    } else {
      logger.info('Using ClientCredentials mode');
      if (!options.clientSecret) {
        logger.error('Client secret missing for ClientCredentials mode');
        throw new Error('Client secret required for ClientCredentials mode');
      }
      const credential = new ClientSecretCredential(options.tenantId, options.clientId, options.clientSecret);
      authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: env.OAUTH_SCOPES?.split(' ') ?? ['https://graph.microsoft.com/.default'],
      });
    }

    this.client = Client.initWithMiddleware({ authProvider });
    logger.info('MSGraphService client initialized');
  }

  async genericGraphRequest(
    path: string,
    method: string,
    body?: any,
    queryParams?: Record<string, string>,
    version: 'v1.0' | 'beta' = 'v1.0',
    fetchAll = false,
    consistencyLevel?: string
  ): Promise<any> {
    logger.info('Executing genericGraphRequest', { path, method, version, fetchAll, consistencyLevel });
    try {
      const request = this.client.api(path).version(version);
      if (queryParams) {
        request.query(queryParams);
      }
      if (consistencyLevel) {
        request.header('ConsistencyLevel', consistencyLevel);
      }

      let response;
      switch (method.toLowerCase()) {
        case 'get':
          response = await request.get();
          break;
        case 'post':
          response = await request.post(body);
          break;
        case 'put':
          response = await request.put(body);
          break;
        case 'patch':
          response = await request.patch(body);
          break;
        case 'delete':
          response = await request.delete();
          break;
        default:
          logger.error('Invalid HTTP method', { method });
          throw new Error(`Invalid HTTP method: ${method}`);
      }

      if (fetchAll && response['@odata.nextLink']) {
        logger.info('Fetching all pages for Graph request', { nextLink: response['@odata.nextLink'] });
        const allResults = response.value ? [...response.value] : [];
        let nextLink = response['@odata.nextLink'];
        while (nextLink) {
          const nextResponse = await this.client.api(nextLink).get();
          if (nextResponse.value) {
            allResults.push(...nextResponse.value);
          }
          nextLink = nextResponse['@odata.nextLink'];
        }
        response.value = allResults;
      }

      logger.info('genericGraphRequest successful', { path, method });
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('genericGraphRequest error', { msg, path, method });
      throw new Error(msg);
    }
  }

  async azureRequest(
    path: string,
    method: string,
    body?: any,
    queryParams?: Record<string, string>,
    apiVersion?: string,
    subscriptionId?: string,
    fetchAll = false
  ): Promise<any> {
    logger.info('Executing azureRequest', { path, method, apiVersion, subscriptionId, fetchAll });
    try {
      let fullPath = subscriptionId ? `/subscriptions/${subscriptionId}${path}` : path;
      if (apiVersion) {
        fullPath = `${fullPath}${fullPath.includes('?') ? '&' : '?'}api-version=${apiVersion}`;
      }
      const request = this.client.api(fullPath);
      if (queryParams) {
        request.query(queryParams);
      }

      let response;
      switch (method.toLowerCase()) {
        case 'get':
          response = await request.get();
          break;
        case 'post':
          response = await request.post(body);
          break;
        case 'put':
          response = await request.put(body);
          break;
        case 'patch':
          response = await request.patch(body);
          break;
        case 'delete':
          response = await request.delete();
          break;
        default:
          logger.error('Invalid HTTP method', { method });
          throw new Error(`Invalid HTTP method: ${method}`);
      }

      if (fetchAll && response['@odata.nextLink']) {
        logger.info('Fetching all pages for Azure request', { nextLink: response['@odata.nextLink'] });
        const allResults = response.value ? [...response.value] : [];
        let nextLink = response['@odata.nextLink'];
        while (nextLink) {
          const nextResponse = await this.client.api(nextLink).get();
          if (nextResponse.value) {
            allResults.push(...nextResponse.value);
          }
          nextLink = nextResponse['@odata.nextLink'];
        }
        response.value = allResults;
      }

      logger.info('azureRequest successful', { path, method });
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('azureRequest error', { msg, path, method });
      throw new Error(msg);
    }
  }

  async getCurrentUserProfile(): Promise<any> {
    logger.info('Executing getCurrentUserProfile');
    try {
      const response = await this.genericGraphRequest('/me', 'get');
      logger.info('getCurrentUserProfile successful');
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('getCurrentUserProfile error', { msg });
      throw new Error(msg);
    }
  }

  async getUsers(queryParams?: Record<string, string>, fetchAll = false): Promise<any> {
    logger.info('Executing getUsers', { queryParams, fetchAll });
    try {
      const response = await this.genericGraphRequest('/users', 'get', undefined, queryParams, 'v1.0', fetchAll);
      logger.info('getUsers successful');
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('getUsers error', { msg });
      throw new Error(msg);
    }
  }

  async getGroups(queryParams?: Record<string, string>, fetchAll = false): Promise<any> {
    logger.info('Executing getGroups', { queryParams, fetchAll });
    try {
      const response = await this.genericGraphRequest('/groups', 'get', undefined, queryParams, 'v1.0', fetchAll);
      logger.info('getGroups successful');
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('getGroups error', { msg });
      throw new Error(msg);
    }
  }

  async getApplications(queryParams?: Record<string, string>, fetchAll = false): Promise<any> {
    logger.info('Executing getApplications', { queryParams, fetchAll });
    try {
      const response = await this.genericGraphRequest('/applications', 'get', undefined, queryParams, 'v1.0', fetchAll);
      logger.info('getApplications successful');
      return response;
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      logger.error('getApplications error', { msg });
      throw new Error(msg);
    }
  }
}