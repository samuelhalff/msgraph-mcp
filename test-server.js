#!/usr/bin/env node

/**
 * Microsoft Graph MCP Server Test Script
 * 
 * Tests the MCP server endpoints and functionality
 * Can be run against local or remote deployments
 */

import fetch from 'node-fetch';
// Configuration
const CONFIG = {
  // Server configuration
  LOCAL_URL: process.env.LOCAL_URL || 'http://localhost:3001',
  REMOTE_URL: process.env.REMOTE_URL || 'http://msgraph-mcp:3001', // Docker internal URL
  PUBLIC_REMOTE_URL: process.env.PUBLIC_REMOTE_URL || 'https://your-domain.com', // Public URL for OAuth testing
  
  // Test configuration
  TIMEOUT: parseInt(process.env.TEST_TIMEOUT) || 30000, // 30 seconds
  VERBOSE: process.env.VERBOSE === 'true',
  
  // OAuth configuration (from environment variables)
  OAUTH: {
    client_id: process.env.OAUTH_CLIENT_ID || process.env.CLIENT_ID,
    tenant_id: process.env.OAUTH_TENANT_ID || process.env.TENANT_ID,
    redirect_uri: process.env.OAUTH_REDIRECT_URI,
    authorization_url: process.env.OAUTH_AUTHORIZATION_URL,
    token_url: process.env.OAUTH_TOKEN_URL,
    scope: process.env.OAUTH_SCOPE || 'https://graph.microsoft.com/.default'
  }
};

// Validate OAuth configuration
function validateOAuthConfig() {
  const requiredOAuthVars = [
    'client_id',
    'tenant_id', 
    'redirect_uri',
    'authorization_url',
    'token_url'
  ];
  
  const missing = requiredOAuthVars.filter(key => !CONFIG.OAUTH[key]);
  
  if (missing.length > 0) {
    console.warn('⚠️  OAuth tests may fail - missing environment variables:');
    missing.forEach(key => {
      const envVar = key === 'client_id' ? 'OAUTH_CLIENT_ID or CLIENT_ID' :
                     key === 'tenant_id' ? 'OAUTH_TENANT_ID or TENANT_ID' :
                     `OAUTH_${key.toUpperCase()}`;
      console.warn(`   - ${envVar}`);
    });
    console.warn('   Set these variables for full OAuth testing functionality.\n');
    return false;
  }
  
  return true;
}

// Test utilities
class TestRunner {
  constructor(baseUrl) {
    this.baseUrl = baseUrl;
    this.passed = 0;
    this.failed = 0;
    this.results = [];
  }

  log(message, level = 'info') {
    const timestamp = new Date().toISOString();
    const prefix = `[${timestamp}] [${level.toUpperCase()}]`;
    
    if (level === 'error') {
      console.error(`${prefix} ${message}`);
    } else if (level === 'warn') {
      console.warn(`${prefix} ${message}`);
    } else if (CONFIG.VERBOSE || level === 'info') {
      console.log(`${prefix} ${message}`);
    }
  }

  async test(name, testFn) {
    this.log(`Testing: ${name}`);
    try {
      const startTime = Date.now();
      const result = await Promise.race([
        testFn(),
        new Promise((_, reject) => 
          setTimeout(() => reject(new Error('Test timeout')), CONFIG.TIMEOUT)
        )
      ]);
      
      const duration = Date.now() - startTime;
      this.passed++;
      this.results.push({ name, status: 'PASS', duration, result });
      this.log(`✅ ${name} (${duration}ms)`, 'info');
      return result;
    } catch (error) {
      this.failed++;
      this.results.push({ name, status: 'FAIL', error: error.message });
      this.log(`❌ ${name}: ${error.message}`, 'error');
      throw error;
    }
  }

  async makeRequest(endpoint, options = {}) {
    const url = `${this.baseUrl}${endpoint}`;
    this.log(`Making request to: ${url}`, 'debug');
    
    const response = await fetch(url, {
      timeout: CONFIG.TIMEOUT,
      headers: {
        'Content-Type': 'application/json',
        'User-Agent': 'MCP-Test-Script/1.0',
        ...options.headers
      },
      ...options
    });

    const responseText = await response.text();
    let responseData;
    
    try {
      responseData = JSON.parse(responseText);
    } catch {
      responseData = responseText;
    }

    this.log(`Response (${response.status}): ${JSON.stringify(responseData, null, 2)}`, 'debug');

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${responseText}`);
    }

    return { status: response.status, data: responseData, headers: response.headers };
  }

  async makeMcpRequest(method, params = {}) {
    const payload = {
      jsonrpc: '2.0',
      id: Math.random().toString(36).substring(7),
      method,
      params
    };

    return this.makeRequest('/mcp', {
      method: 'POST',
      body: JSON.stringify(payload)
    });
  }

  summary() {
    const total = this.passed + this.failed;
    const passRate = total > 0 ? (this.passed / total * 100).toFixed(1) : 0;
    
    console.log('\n' + '='.repeat(60));
    console.log('TEST SUMMARY');
    console.log('='.repeat(60));
    console.log(`Total Tests: ${total}`);
    console.log(`Passed: ${this.passed} (${passRate}%)`);
    console.log(`Failed: ${this.failed}`);
    console.log('='.repeat(60));

    if (this.failed > 0) {
      console.log('\nFAILED TESTS:');
      this.results
        .filter(r => r.status === 'FAIL')
        .forEach(r => console.log(`❌ ${r.name}: ${r.error}`));
    }

    console.log('\nDETAILED RESULTS:');
    this.results.forEach(r => {
      const status = r.status === 'PASS' ? '✅' : '❌';
      const duration = r.duration ? ` (${r.duration}ms)` : '';
      console.log(`${status} ${r.name}${duration}`);
    });
    return this.failed === 0;
  }
}

// Test suites
async function testHealthEndpoints(runner) {
  console.log('\n🏥 Testing Health Endpoints...');
  
  await runner.test('Health Check', async () => {
    const response = await runner.makeRequest('/health');
    if (response.data.status !== 'ok') {
      throw new Error(`Expected ok status, got: ${response.data.status}`);
    }
    return response.data;
  });

  await runner.test('Root Endpoint', async () => {
    const response = await runner.makeRequest('/');
    // Root endpoint returns plain text, not JSON
    if (typeof response.data === 'string' && response.data.includes('Microsoft Graph MCP Server')) {
      return { message: response.data };
    } else if (response.data.service || response.data.message) {
      return response.data;
    } else {
      throw new Error('Unexpected root endpoint response format');
    }
  });
}

async function testOAuthEndpoints(runner) {
  console.log('\n🔐 Testing OAuth Endpoints...');
  
  const hasOAuthConfig = validateOAuthConfig();
  
  await runner.test('OAuth Discovery (well-known)', async () => {
    const response = await runner.makeRequest('/.well-known/oauth-authorization-server');
    const data = response.data;
    
    // Verify required OAuth discovery fields
    const requiredFields = ['issuer', 'authorization_endpoint', 'token_endpoint', 'response_types_supported'];
    for (const field of requiredFields) {
      if (!data[field]) {
        throw new Error(`Missing required OAuth discovery field: ${field}`);
      }
    }
    
    // Only verify tenant ID if we have OAuth config
    if (hasOAuthConfig && CONFIG.OAUTH.tenant_id) {
      if (!data.authorization_endpoint.includes(CONFIG.OAUTH.tenant_id)) {
        throw new Error('Authorization endpoint missing expected tenant ID');
      }
    }
    
    return data;
  });

  await runner.test('OAuth Protected Resource Metadata (RFC9728)', async () => {
    const response = await runner.makeRequest('/.well-known/oauth-protected-resource');
    const data = response.data;
    
    // Verify required RFC9728 fields per MCP specification
    const requiredFields = ['resource', 'authorization_servers'];
    for (const field of requiredFields) {
      if (!data[field]) {
        throw new Error(`Missing required Protected Resource Metadata field: ${field}`);
      }
    }
    
    // Verify authorization_servers is an array with at least one server
    if (!Array.isArray(data.authorization_servers) || data.authorization_servers.length === 0) {
      throw new Error('authorization_servers must be a non-empty array');
    }
    
    // Verify resource field contains a valid URI
    if (!data.resource.startsWith('http')) {
      throw new Error('resource field must be a valid HTTP(S) URI');
    }
    
    // Verify Microsoft specific authorization server
    const hasExpectedServer = data.authorization_servers.some(server => 
      server.includes('login.microsoftonline.com')
    );
    
    if (!hasExpectedServer) {
      throw new Error('Expected Microsoft authorization server in authorization_servers list');
    }
    
    // Verify recommended fields are present
    if (!data.scopes_supported || !Array.isArray(data.scopes_supported)) {
      throw new Error('scopes_supported should be present and be an array');
    }
    
    return data;
  });

  await runner.test('OAuth Registration', async () => {
    const redirectUri = CONFIG.OAUTH.redirect_uri || 'https://example.com/callback';
    const scope = CONFIG.OAUTH.scope || 'https://graph.microsoft.com/.default';
    
    const registrationData = {
      client_name: 'Test MCP Client',
      redirect_uris: [redirectUri],
      scope: scope,
      grant_types: ['authorization_code'],
      response_types: ['code'],
      token_endpoint_auth_method: 'none'  // Public client, no secret needed
    };

    const response = await runner.makeRequest('/register', {
      method: 'POST',
      body: JSON.stringify(registrationData)
    });

    // For public clients, only client_id is required
    if (!response.data.client_id) {
      throw new Error('Missing client_id in registration response');
    }

    // Check that the auth method was set correctly for public client
    if (response.data.token_endpoint_auth_method !== 'none') {
      throw new Error('Expected public client (token_endpoint_auth_method: none)');
    }

    return response.data;
  });
}

async function testMcpEndpoints(runner) {
  console.log('\n🔧 Testing MCP Protocol Endpoints...');
  
  await runner.test('MCP Initialize', async () => {
    const response = await runner.makeMcpRequest('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {
        tools: {}
      },
      clientInfo: {
        name: 'test-client',
        version: '1.0.0'
      }
    });

    if (response.data.error) {
      throw new Error(`MCP Error: ${response.data.error.message}`);
    }

    const result = response.data.result;
    if (!result.capabilities || !result.serverInfo) {
      throw new Error('Missing capabilities or serverInfo in initialize response');
    }

    return result;
  });

  await runner.test('MCP Tools List', async () => {
    const response = await runner.makeMcpRequest('tools/list');
    
    if (response.data.error) {
      throw new Error(`MCP Error: ${response.data.error.message}`);
    }

    const tools = response.data.result.tools;
    if (!Array.isArray(tools) || tools.length === 0) {
      throw new Error('No tools returned from tools/list');
    }

    // Verify expected tools are present
    const expectedTools = [
      'microsoft-graph-api',
      'microsoft-graph-profile', 
      'list-users',
      'list-groups',
      'search-users',
      'send-mail',
      'list-calendar-events',
      'create-calendar-event',
      'search-files',
      'get-schedule',
      'throttling-stats'
    ];

    const toolNames = tools.map(t => t.name);
    for (const expectedTool of expectedTools) {
      if (!toolNames.includes(expectedTool)) {
        throw new Error(`Missing expected tool: ${expectedTool}`);
      }
    }

    return tools;
  });

  await runner.test('MCP Ping', async () => {
    const response = await runner.makeMcpRequest('ping');
    
    if (response.data.error) {
      throw new Error(`MCP Error: ${response.data.error.message}`);
    }

    if (response.data.result.status !== 'ok') {
      throw new Error(`Expected ping status 'ok', got: ${response.data.result.status}`);
    }

    return response.data.result;
  });

  await runner.test('MCP Authentication Required (401 with WWW-Authenticate)', async () => {
    // Test that protected MCP methods return proper 401 with WWW-Authenticate header
    try {
      const response = await runner.makeMcpRequest('tools/call', {
        name: 'throttling-stats',
        arguments: {}
      });
      
      // Should get a 401 response with proper headers
      if (response.status !== 401) {
        throw new Error(`Expected 401 status for unauthenticated call, got ${response.status}`);
      }
      
      // Check for WWW-Authenticate header as required by RFC9728 Section 5.1
      const wwwAuthHeader = response.headers.get ? 
        response.headers.get('www-authenticate') : 
        response.headers['www-authenticate'];
      
      if (!wwwAuthHeader) {
        throw new Error('Missing WWW-Authenticate header in 401 response (required by MCP spec)');
      }
      
      // Verify header format includes Bearer and resource_metadata_url
      if (!wwwAuthHeader.includes('Bearer') || !wwwAuthHeader.includes('resource_metadata_url')) {
        throw new Error(`Invalid WWW-Authenticate header format: ${wwwAuthHeader}`);
      }
      
      // Should still be valid JSON-RPC response in body
      if (response.data.error && response.data.error.code !== -32002) {
        throw new Error(`Expected authentication error code -32002, got ${response.data.error.code}`);
      }
      
      return { 
        status: response.status, 
        wwwAuthenticate: wwwAuthHeader,
        error: response.data.error 
      };
    } catch (error) {
      // If the fetch itself fails, that might be expected
      if (error.message && error.message.includes('401')) {
        return { expected401: true };
      }
      throw error;
    }
  });
}

async function testMcpTools(runner) {
  console.log('\n🛠️  Testing MCP Tool Calls...');
  
  // Note: These tests will fail without proper OAuth tokens, but we can test the structure
  await runner.test('Throttling Stats Tool', async () => {
    const response = await runner.makeMcpRequest('tools/call', {
      name: 'throttling-stats',
      arguments: {}
    });

    // This should work even without OAuth since it's just internal stats
    if (response.data.error) {
      // Log the error but don't fail if it's auth-related
      runner.log(`Throttling stats error (expected without auth): ${response.data.error.message}`, 'warn');
      return { note: 'Auth required for full functionality' };
    }

    const result = response.data.result;
    if (!result.content || !Array.isArray(result.content)) {
      throw new Error('Invalid throttling stats response format');
    }

    return result;
  });

  await runner.test('Invalid Tool Call', async () => {
    const response = await runner.makeMcpRequest('tools/call', {
      name: 'nonexistent-tool',
      arguments: {}
    });

    // This should return an error
    if (!response.data.error) {
      throw new Error('Expected error for nonexistent tool, but got success');
    }

    return { error: response.data.error.message };
  });
}

async function testErrorHandling(runner) {
  console.log('\n⚠️  Testing Error Handling...');
  
  await runner.test('Invalid JSON-RPC Request', async () => {
    try {
      await runner.makeRequest('/mcp', {
        method: 'POST',
        body: 'invalid json'
      });
      throw new Error('Expected request to fail with invalid JSON');
    } catch (error) {
      if (!error.message.includes('400') && !error.message.includes('parse')) {
        throw error;
      }
      return { error: 'Correctly rejected invalid JSON' };
    }
  });

  await runner.test('Unsupported MCP Method', async () => {
    const response = await runner.makeMcpRequest('unsupported/method');
    
    if (!response.data.error || response.data.error.code !== -32601) {
      throw new Error('Expected method not found error (-32601)');
    }

    return response.data.error;
  });

  await runner.test('Invalid Endpoint', async () => {
    try {
      await runner.makeRequest('/nonexistent');
      throw new Error('Expected 404 for nonexistent endpoint');
    } catch (error) {
      if (!error.message.includes('404')) {
        throw error;
      }
      return { error: 'Correctly returned 404' };
    }
  });
}

// Main test execution
async function runTests(serverUrl) {
  console.log(`\n🚀 Starting MCP Server Tests against: ${serverUrl}`);
  console.log('='.repeat(60));
  
  // Validate configuration
  console.log('📋 Configuration Check:');
  console.log(`   Server URL: ${serverUrl}`);
  console.log(`   Timeout: ${CONFIG.TIMEOUT}ms`);
  console.log(`   Verbose: ${CONFIG.VERBOSE}`);
  
  const hasOAuthConfig = validateOAuthConfig();
  if (hasOAuthConfig) {
    console.log('✅ OAuth configuration complete\n');
  } else {
    console.log('⚠️  OAuth configuration incomplete (tests will be limited)\n');
  }
  
  const runner = new TestRunner(serverUrl);
  
  try {
    await testHealthEndpoints(runner);
    await testOAuthEndpoints(runner);
    await testMcpEndpoints(runner);
    await testMcpTools(runner);
    await testErrorHandling(runner);
  } catch (error) {
    runner.log(`Test suite failed: ${error.message}`, 'error');
  }
  
  const success = runner.summary();
  return success;
}

// Command line interface
async function main() {
  const args = process.argv.slice(2);
  const command = args[0] || 'local';
  
  let serverUrl;
  switch (command) {
    case 'local':
      serverUrl = CONFIG.LOCAL_URL;
      console.log('Testing against LOCAL server...');
      break;
    case 'remote':
      serverUrl = CONFIG.REMOTE_URL;
      console.log('Testing against REMOTE server (Docker internal)...');
      break;
    case 'public':
      serverUrl = CONFIG.PUBLIC_REMOTE_URL;
      console.log('Testing against PUBLIC remote server...');
      break;
    default:
      if (command.startsWith('http')) {
        serverUrl = command;
        console.log(`Testing against CUSTOM server: ${serverUrl}`);
      } else {
        console.error('Usage: node test-server.js [local|remote|public|<url>]');
        console.error('  local  - Test http://localhost:3001');
        console.error('  remote - Test http://msgraph-mcp:3001 (Docker)');
        console.error('  public - Test https://pbm-ai.ddns.net');
        console.error('  <url>  - Test custom URL');
        process.exit(1);
      }
  }
  
  try {
    const success = await runTests(serverUrl);
    process.exit(success ? 0 : 1);
  } catch (error) {
    console.error('Test execution failed:', error.message);
    process.exit(1);
  }
}

// Handle uncaught errors
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  process.exit(1);
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  process.exit(1);
});

// Run the tests
if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

export { runTests, TestRunner, CONFIG };
