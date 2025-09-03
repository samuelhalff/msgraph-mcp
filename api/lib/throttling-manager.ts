/**
 * Microsoft Graph Throttling Manager
 * 
 * Implements throttling limits and retry logic for Microsoft Graph API calls
 * Based on: https://learn.microsoft.com/en-us/graph/throttling-limits
 */

import logger from './logger.js';

export interface ThrottlingHeaders {
  'retry-after'?: string;
  'ratelimit-limit'?: string;
  'ratelimit-remaining'?: string;
  'ratelimit-reset'?: string;
}

export interface RequestMetrics {
  timestamp: number;
  endpoint: string;
  method: string;
  statusCode?: number;
  retryAfter?: number;
}

export class ThrottlingManager {
  private requestHistory: RequestMetrics[] = [];
  private readonly maxHistorySize = 1000;
  
  // Throttling limits per service (requests per second)
  private readonly serviceLimits = {
    // Microsoft Graph core services
    'graph.microsoft.com': {
      default: 10000, // 10,000 requests per 10 minutes per app
      '/me': 2000,    // User-specific endpoints
      '/users': 1600, // User operations
      '/groups': 1600, // Group operations
      '/sites': 3200, // SharePoint sites
      '/drives': 3200, // OneDrive/SharePoint drives
      '/search': 500,  // Search API (more restrictive)
      '/teams': 1000,  // Teams API
      '/calendar': 1000, // Calendar API
      '/mail': 1000,   // Mail API
    }
  };

  // Exponential backoff configuration
  private readonly retryConfig = {
    maxRetries: 3,
    baseDelayMs: 1000,
    maxDelayMs: 60000,
    backoffMultiplier: 2,
    jitterFactor: 0.1
  };

  /**
   * Check if a request should be allowed based on current throttling state
   */
  shouldAllowRequest(endpoint: string, method: string = 'GET'): boolean {
    const now = Date.now();
    const windowMs = 600000; // 10 minutes window
    
    // Clean old entries
    this.requestHistory = this.requestHistory.filter(
      req => now - req.timestamp < windowMs
    );

    // Get service-specific limit
    const limit = this.getEndpointLimit(endpoint);
    const recentRequests = this.requestHistory.filter(
      req => req.endpoint === endpoint && req.method === method
    ).length;

    logger.debug('Throttling check', {
      endpoint,
      method,
      recentRequests,
      limit,
      allowed: recentRequests < limit
    });

    return recentRequests < limit;
  }

  /**
   * Record a request for throttling tracking
   */
  recordRequest(endpoint: string, method: string, statusCode?: number, headers?: Record<string, string>): void {
    const metrics: RequestMetrics = {
      timestamp: Date.now(),
      endpoint,
      method,
      statusCode
    };

    // Extract retry-after header if present
    if (headers?.['retry-after']) {
      metrics.retryAfter = parseInt(headers['retry-after'], 10);
    }

    this.requestHistory.push(metrics);

    // Maintain history size
    if (this.requestHistory.length > this.maxHistorySize) {
      this.requestHistory = this.requestHistory.slice(-this.maxHistorySize);
    }

    logger.debug('Request recorded', { metrics });
  }

  /**
   * Calculate delay before retry based on response headers and attempt number
   */
  calculateRetryDelay(attempt: number, headers?: Record<string, string>): number {
    // Use Retry-After header if provided
    if (headers?.['retry-after']) {
      const retryAfter = parseInt(headers['retry-after'], 10);
      if (!isNaN(retryAfter)) {
        return retryAfter * 1000; // Convert to milliseconds
      }
    }

    // Exponential backoff with jitter
    const baseDelay = this.retryConfig.baseDelayMs * Math.pow(this.retryConfig.backoffMultiplier, attempt - 1);
    const jitter = baseDelay * this.retryConfig.jitterFactor * Math.random();
    const delay = Math.min(baseDelay + jitter, this.retryConfig.maxDelayMs);

    logger.debug('Calculated retry delay', {
      attempt,
      baseDelay,
      jitter,
      finalDelay: delay,
      hasRetryAfter: !!headers?.['retry-after']
    });

    return delay;
  }

  /**
   * Execute a request with automatic retry logic for throttling
   */
  async executeWithRetry<T>(
    requestFn: () => Promise<{ data: T; status: number; headers: Record<string, string> }>,
    endpoint: string,
    method: string = 'GET'
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= this.retryConfig.maxRetries + 1; attempt++) {
      try {
        // Check throttling before request (except for retries after 429)
        if (attempt === 1 && !this.shouldAllowRequest(endpoint, method)) {
          throw new Error(`Request rate limit exceeded for ${endpoint}`);
        }

        logger.debug('Executing request', { endpoint, method, attempt });

        const response = await requestFn();
        
        // Record successful request
        this.recordRequest(endpoint, method, response.status, response.headers);

        return response.data;

      } catch (error) {
        lastError = error as Error;
        
        // Extract status code from error
        const statusCode = this.extractStatusCode(error);
        const headers = this.extractHeaders(error);

        // Record failed request
        this.recordRequest(endpoint, method, statusCode, headers);

        logger.warn('Request failed', {
          endpoint,
          method,
          attempt,
          statusCode,
          error: lastError.message
        });

        // Only retry on throttling errors (429) or server errors (5xx)
        const shouldRetry = (statusCode === 429 || (statusCode !== undefined && statusCode >= 500 && statusCode < 600)) 
                           && attempt <= this.retryConfig.maxRetries;

        if (!shouldRetry) {
          throw lastError;
        }

        // Calculate delay and wait
        const delay = this.calculateRetryDelay(attempt, headers);
        
        logger.info('Retrying request after delay', {
          endpoint,
          method,
          attempt,
          delayMs: delay,
          statusCode
        });

        await this.sleep(delay);
      }
    }

    throw lastError || new Error('Max retries exceeded');
  }

  /**
   * Get throttling statistics for monitoring
   */
  getStats(): {
    totalRequests: number;
    recentRequests: number;
    errorRate: number;
    throttledRequests: number;
  } {
    const now = Date.now();
    const windowMs = 600000; // 10 minutes

    const recentRequests = this.requestHistory.filter(
      req => now - req.timestamp < windowMs
    );

    const errorRequests = recentRequests.filter(
      req => req.statusCode && req.statusCode >= 400
    );

    const throttledRequests = recentRequests.filter(
      req => req.statusCode === 429
    );

    return {
      totalRequests: this.requestHistory.length,
      recentRequests: recentRequests.length,
      errorRate: recentRequests.length > 0 ? errorRequests.length / recentRequests.length : 0,
      throttledRequests: throttledRequests.length
    };
  }

  /**
   * Reset throttling state (useful for testing)
   */
  reset(): void {
    this.requestHistory = [];
    logger.debug('Throttling manager reset');
  }

  // Private helper methods

  private getEndpointLimit(endpoint: string): number {
    const graphLimits = this.serviceLimits['graph.microsoft.com'];
    
    // Find most specific matching limit
    const matchingKeys = Object.keys(graphLimits)
      .filter(key => key !== 'default' && endpoint.includes(key))
      .sort((a, b) => b.length - a.length); // Sort by specificity (longer first)

    return matchingKeys.length > 0 
      ? graphLimits[matchingKeys[0] as keyof typeof graphLimits] as number
      : graphLimits.default;
  }

  private extractStatusCode(error: unknown): number | undefined {
    if (typeof error === 'object' && error !== null) {
      const err = error as Record<string, unknown>;
      const response = err.response as Record<string, unknown>;
      return (err.status as number) || (err.statusCode as number) || (response?.status as number);
    }
    return undefined;
  }

  private extractHeaders(error: unknown): Record<string, string> | undefined {
    if (typeof error === 'object' && error !== null) {
      const err = error as Record<string, unknown>;
      const response = err.response as Record<string, unknown>;
      return (err.headers as Record<string, string>) || (response?.headers as Record<string, string>);
    }
    return undefined;
  }

  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Singleton instance for global use
export const throttlingManager = new ThrottlingManager();
