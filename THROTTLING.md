# Microsoft Graph Throttling Implementation

This document describes the throttling implementation for the Microsoft Graph MCP server, based on [Microsoft Graph throttling limits](https://learn.microsoft.com/en-us/graph/throttling-limits).

## Overview

The throttling manager implements Microsoft Graph's best practices for handling API rate limits and provides automatic retry logic with exponential backoff.

## Features

### 1. Rate Limit Tracking
- Tracks request history over a 10-minute sliding window
- Monitors requests per endpoint with service-specific limits
- Provides real-time throttling statistics

### 2. Service-Specific Limits
The implementation includes specific limits for different Microsoft Graph services:

| Service | Requests per 10 minutes |
|---------|------------------------|
| Default | 10,000 |
| User endpoints (`/me`) | 2,000 |
| Users (`/users`) | 1,600 |
| Groups (`/groups`) | 1,600 |
| SharePoint Sites (`/sites`) | 3,200 |
| OneDrive/SharePoint (`/drives`) | 3,200 |
| Search API (`/search`) | 500 |
| Teams API (`/teams`) | 1,000 |
| Calendar API (`/calendar`) | 1,000 |
| Mail API (`/mail`) | 1,000 |

### 3. Automatic Retry Logic
- **Retry Conditions**: Automatically retries on HTTP 429 (throttled) and 5xx (server errors)
- **Max Retries**: 3 attempts per request
- **Exponential Backoff**: Base delay of 1 second, multiplied by 2 for each retry
- **Jitter**: Random 10% variance to prevent thundering herd
- **Retry-After Header**: Respects Microsoft's retry-after headers when provided
- **Max Delay**: Capped at 60 seconds

### 4. Request Monitoring
The throttling manager tracks:
- Total request count
- Recent requests (10-minute window)
- Error rate percentage
- Number of throttled requests (429 responses)

## Usage

### Automatic Integration
All Microsoft Graph requests through `MSGraphService.genericGraphRequest()` automatically use throttling:

```typescript
// This call will automatically handle throttling and retries
const result = await this.svc.genericGraphRequest('/me/profile', 'get');
```

### Manual Throttling Check
```typescript
import { throttlingManager } from './lib/throttling-manager.js';

// Check if request should be allowed
const allowed = throttlingManager.shouldAllowRequest('/me/calendar/events', 'GET');

// Execute with retry logic
const result = await throttlingManager.executeWithRetry(
  async () => {
    // Your request logic here
    return { data: response, status: 200, headers: {} };
  },
  '/me/calendar/events',
  'GET'
);
```

### Monitoring Statistics
Use the `throttling-stats` MCP tool to monitor API usage:

```json
{
  "totalRequests": 1250,
  "recentRequests": 45,
  "errorRate": 0.02,
  "throttledRequests": 1,
  "timestamp": "2024-03-15T10:30:00.000Z",
  "windowSize": "10 minutes"
}
```

## MCP Tools

### throttling-stats
- **Description**: Get current throttling statistics and API performance metrics
- **Parameters**: None
- **Returns**: Statistics object with request counts, error rates, and throttled requests

## Configuration

### Environment Variables
No additional environment variables are required. The throttling manager uses sensible defaults based on Microsoft Graph documentation.

### Customization
To modify throttling limits, edit the `serviceLimits` object in `throttling-manager.ts`:

```typescript
private readonly serviceLimits = {
  'graph.microsoft.com': {
    default: 10000,
    '/custom-endpoint': 500,
    // Add your custom limits here
  }
};
```

## Error Handling

### Throttling Errors (429)
- Automatic retry with exponential backoff
- Respects `Retry-After` headers from Microsoft Graph
- Logs throttling events for monitoring

### Server Errors (5xx)
- Automatic retry for transient server issues
- Same retry logic as throttling errors
- Distinguishes between client and server errors

### Client Errors (4xx)
- No automatic retry (except for 429)
- Immediate failure with error details
- Proper error propagation to MCP client

## Best Practices

### 1. Batch Operations
For bulk operations, consider using:
- Microsoft Graph batch requests (`/$batch`)
- Delta queries for incremental sync
- Pagination with appropriate page sizes

### 2. Monitoring
- Regularly check throttling statistics
- Monitor error rates and adjust usage patterns
- Use consistent headers and user-agent strings

### 3. Efficient Queries
- Use specific field selectors (`$select`)
- Implement proper filtering (`$filter`)
- Cache responses when appropriate

## Troubleshooting

### High Throttling Rates
1. Check `throttling-stats` for current usage
2. Reduce request frequency
3. Implement client-side caching
4. Use more specific API endpoints

### Retry Failures
1. Verify network connectivity
2. Check authentication token validity
3. Review Microsoft Graph service status
4. Increase retry delays for high-load scenarios

## Integration with LibreChat

The throttling manager is transparent to LibreChat. All MCP tool calls automatically benefit from throttling protection:

1. LibreChat sends MCP requests
2. Server validates and routes requests
3. Throttling manager handles rate limiting
4. Microsoft Graph API calls are made safely
5. Responses are returned to LibreChat

## Future Enhancements

### Planned Features
- [ ] Per-user throttling limits
- [ ] Redis-backed distributed throttling
- [ ] Advanced analytics and alerting
- [ ] Dynamic limit adjustment based on quota headers
- [ ] Integration with Azure Monitor metrics

### Potential Optimizations
- [ ] Request queuing for burst handling
- [ ] Predictive throttling based on usage patterns
- [ ] Integration with Microsoft Graph webhooks
- [ ] Custom retry strategies per endpoint type
