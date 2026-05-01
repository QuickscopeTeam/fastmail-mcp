import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { createMcpExpressApp } from '@modelcontextprotocol/sdk/server/express.js';

const PORT = 3001;

// Create Express app with MCP middleware
const app = createMcpExpressApp();

// Create a proxy MCP server that forwards to the stdio Fastmail MCP
const createProxyServer = async () => {
  const server = new McpServer(
    { name: 'fastmail-mcp-http-proxy', version: '1.0.0' },
    { capabilities: { tools: {} } }
  );

  // Create client to connect to stdio MCP server
  const client = new Client(
    { name: 'http-proxy-client', version: '1.0.0' },
    { capabilities: {} }
  );

  const transport = new StdioClientTransport({
    command: 'node',
    args: ['/Users/ryanpage/fastmail-mcp/index.js'],
    env: process.env
  });

  await client.connect(transport);

  // Forward tools/list requests
  server.setRequestHandler({ method: 'tools/list' }, async () => {
    const tools = await client.listTools();
    return tools;
  });

  // Forward tools/call requests
  server.setRequestHandler({ method: 'tools/call' }, async (request) => {
    const result = await client.callTool(request.params);
    return result;
  });

  return server;
};

// Handle POST requests to /mcp endpoint
app.post('/mcp', async (req, res) => {
  console.log('Received POST MCP request');
  
  try {
    const server = await createProxyServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => Math.random().toString(36).substring(2, 15),
      allowedHosts: ['localhost:3001', 'cure-education-prism.ngrok-free.dev', '127.0.0.1:3001'],
      enableDnsRebindingProtection: false
    });
    
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
    
    res.on('close', () => {
      console.log('Request closed');
      transport.close();
      server.close();
    });
  } catch (error) {
    console.error('Error handling MCP request:', error);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: {
          code: -32603,
          message: 'Internal server error: ' + error.message
        },
        id: null
      });
    }
  }
});

// Handle GET requests to /mcp endpoint
app.get('/mcp', async (req, res) => {
  console.log('Received GET MCP request');
  res.writeHead(405).end(JSON.stringify({
    jsonrpc: '2.0',
    error: {
      code: -32000,
      message: 'Method not allowed. Use POST.'
    },
    id: null
  }));
});

// Start the HTTP server
app.listen(PORT, (error) => {
  if (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
  console.log(`Fastmail MCP HTTP Proxy (Streamable) running on port ${PORT}`);
  console.log(`Endpoint: http://localhost:${PORT}/mcp`);
});

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.log('Shutting down HTTP proxy...');
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('Shutting down HTTP proxy...');
  process.exit(0);
});
