#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import fetch from "node-fetch";
import http from "http";
import url from "url";

const FASTMAIL_SESSION_URL = "https://api.fastmail.com/.well-known/jmap";
const API_KEY = process.env.FASTMAIL_API_KEY || "fmu1-d01e43a8-f3ca5b7579eb5aac1f2df46b23440060-0-e12f16ead889af9da7e5dbf2720a49ff";
const PORT = process.env.PORT || 3000;

// Shared session and account ID cache
let cachedSession = null;
let cachedAccountId = null;

async function getSession() {
  if (cachedSession) return cachedSession;

  const response = await fetch(FASTMAIL_SESSION_URL, {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${API_KEY}`,
    },
  });

  if (!response.ok) {
    throw new Error(`JMAP session request failed: ${response.statusText}`);
  }

  cachedSession = await response.json();
  return cachedSession;
}

async function getAccountId() {
  if (cachedAccountId) return cachedAccountId;

  const session = await getSession();
  cachedAccountId = session.primaryAccounts["urn:ietf:params:jmap:mail"];
  return cachedAccountId;
}

async function jmapRequest(methodCalls) {
  const session = await getSession();
  const response = await fetch(session.apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${API_KEY}`,
    },
    body: JSON.stringify({
      using: [
        "urn:ietf:params:jmap:core",
        "urn:ietf:params:jmap:mail",
        "urn:ietf:params:jmap:submission",
      ],
      methodCalls,
    }),
  });

  if (!response.ok) {
    throw new Error(`JMAP request failed: ${response.statusText}`);
  }

  return await response.json();
}

// Tool implementation functions
async function listInbox(accountId, limit) {
  // First get the inbox mailbox ID
  const mailboxResponse = await jmapRequest([
    [
      "Mailbox/get",
      {
        accountId,
      },
      "0",
    ],
  ]);

  const inboxMailbox = mailboxResponse.methodResponses[0][1].list.find(
    (mb) => mb.role === "inbox"
  );

  if (!inboxMailbox) {
    throw new Error("Inbox mailbox not found");
  }

  // Then query emails in the inbox
  const response = await jmapRequest([
    [
      "Email/query",
      {
        accountId,
        filter: { inMailbox: inboxMailbox.id },
        sort: [{ property: "receivedAt", isAscending: false }],
        limit,
      },
      "0",
    ],
    [
      "Email/get",
      {
        accountId,
        "#ids": {
          resultOf: "0",
          name: "Email/query",
          path: "/ids",
        },
        properties: ["id", "subject", "from", "receivedAt", "preview"],
      },
      "1",
    ],
  ]);

  const emails = response.methodResponses[1][1].list;
  const formatted = emails.map((email) => ({
    id: email.id,
    subject: email.subject,
    from: email.from?.[0]?.email || "unknown",
    date: email.receivedAt,
    preview: email.preview,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(formatted, null, 2),
      },
    ],
  };
}

async function searchEmail(accountId, args) {
  const filter = {};
  if (args.query) filter.text = args.query;
  if (args.sender) filter.from = args.sender;
  if (args.subject) filter.subject = args.subject;

  const response = await jmapRequest([
    [
      "Email/query",
      {
        accountId,
        filter: Object.keys(filter).length > 0 ? filter : null,
        sort: [{ property: "receivedAt", isAscending: false }],
        limit: args.limit || 20,
      },
      "0",
    ],
    [
      "Email/get",
      {
        accountId,
        "#ids": {
          resultOf: "0",
          name: "Email/query",
          path: "/ids",
        },
        properties: ["id", "subject", "from", "receivedAt", "preview"],
      },
      "1",
    ],
  ]);

  const emails = response.methodResponses[1][1].list;
  const formatted = emails.map((email) => ({
    id: email.id,
    subject: email.subject,
    from: email.from?.[0]?.email || "unknown",
    date: email.receivedAt,
    preview: email.preview,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(formatted, null, 2),
      },
    ],
  };
}

async function readEmail(accountId, emailId) {
  const response = await jmapRequest([
    [
      "Email/get",
      {
        accountId,
        ids: [emailId],
        properties: [
          "id",
          "subject",
          "from",
          "to",
          "cc",
          "bcc",
          "receivedAt",
          "textBody",
          "htmlBody",
          "bodyValues",
        ],
      },
      "0",
    ],
  ]);

  const email = response.methodResponses[0][1].list[0];
  if (!email) {
    throw new Error("Email not found");
  }

  const textBodyPartId = email.textBody?.[0]?.partId;
  const htmlBodyPartId = email.htmlBody?.[0]?.partId;
  const bodyContent = textBodyPartId
    ? email.bodyValues[textBodyPartId]?.value
    : htmlBodyPartId
    ? email.bodyValues[htmlBodyPartId]?.value
    : "No body content";

  const formatted = {
    id: email.id,
    subject: email.subject,
    from: email.from,
    to: email.to,
    cc: email.cc,
    bcc: email.bcc,
    date: email.receivedAt,
    body: bodyContent,
  };

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(formatted, null, 2),
      },
    ],
  };
}

async function sendEmail(accountId, args) {
  const emailId = `draft-${Date.now()}`;
  
  // Get identity if fromAlias is specified
  const identityResponse = args.fromAlias 
    ? await jmapRequest([["Identity/get", { accountId, ids: [args.fromAlias] }, "0"]])
    : null;

  // Get the from email address if fromAlias is specified
  let fromEmail = undefined;
  let identityId = null;
  if (args.fromAlias && identityResponse) {
    const identity = identityResponse.methodResponses[0][1].list?.[0];
    if (!identity) {
      throw new Error(`Identity ${args.fromAlias} not found`);
    }
    fromEmail = [{ email: identity.email }];
    identityId = args.fromAlias;
  }

  // Ensure body has proper HTML structure
  let emailBody = args.body;
  if (!emailBody.trim().toLowerCase().startsWith('<html')) {
    emailBody = `<html><body style="font-family: sans-serif;">${emailBody}</body></html>`;
  }

  // Create and send email (no mailbox needed for transient submission)
  const response = await jmapRequest([
    [
      "Email/set",
      {
        accountId,
        create: {
          [emailId]: {
            mailboxIds: {},
            from: fromEmail,
            to: args.to.map((email) => ({ email })),
            subject: args.subject,
            htmlBody: [{ partId: "body", type: "text/html" }],
            bodyValues: {
              body: { value: emailBody, charset: "utf-8", isTruncated: false },
            },
          },
        },
      },
      "0",
    ],
    [
      "EmailSubmission/set",
      {
        accountId,
        onSuccessDestroyEmail: [`#${emailId}`],
        create: {
          submission1: {
            emailId: `#${emailId}`,
            identityId,
          },
        },
      },
      "1",
    ],
  ]);

  // Check for submission errors
  const submissionResult = response.methodResponses[1][1];
  if (submissionResult.notCreated) {
    const error = submissionResult.notCreated.submission1;
    throw new Error(`Email submission failed: ${error.description || JSON.stringify(error)}`);
  }

  return {
    content: [
      {
        type: "text",
        text: "Email sent successfully",
      },
    ],
  };
}

async function listAliases(accountId) {
  const response = await jmapRequest([
    [
      "Identity/get",
      {
        accountId,
      },
      "0",
    ],
  ]);

  const identities = response.methodResponses[0][1].list;
  
  // Log raw identity data to verify what Fastmail returns
  console.log('[listAliases] Raw identities:', JSON.stringify(identities, null, 2));
  
  const formatted = identities.map((identity) => ({
    id: identity.id,
    name: identity.name,
    email: identity.email,
    textSignature: identity.textSignature || "",
    htmlSignature: identity.htmlSignature || "",
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(formatted, null, 2),
      },
    ],
  };
}

async function listFolders(accountId) {
  const response = await jmapRequest([
    [
      "Mailbox/get",
      {
        accountId,
      },
      "0",
    ],
  ]);

  const mailboxes = response.methodResponses[0][1].list;
  const formatted = mailboxes.map((mailbox) => ({
    id: mailbox.id,
    name: mailbox.name,
    role: mailbox.role,
    totalEmails: mailbox.totalEmails,
    unreadEmails: mailbox.unreadEmails,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(formatted, null, 2),
      },
    ],
  };
}

async function moveEmail(accountId, emailId, mailboxId) {
  await jmapRequest([
    [
      "Email/set",
      {
        accountId,
        update: {
          [emailId]: {
            mailboxIds: { [mailboxId]: true },
          },
        },
      },
      "0",
    ],
  ]);

  return {
    content: [
      {
        type: "text",
        text: "Email moved successfully",
      },
    ],
  };
}

async function deleteEmail(accountId, emailId) {
  const mailboxResponse = await jmapRequest([
    [
      "Mailbox/get",
      {
        accountId,
      },
      "0",
    ],
  ]);

  const trashMailbox = mailboxResponse.methodResponses[0][1].list.find(
    (mb) => mb.role === "trash"
  );

  if (!trashMailbox) {
    throw new Error("Trash mailbox not found");
  }

  return await moveEmail(accountId, emailId, trashMailbox.id);
}

async function markRead(accountId, emailId, isRead) {
  await jmapRequest([
    [
      "Email/set",
      {
        accountId,
        update: {
          [emailId]: {
            keywords: isRead ? { $seen: true } : {},
          },
        },
      },
      "0",
    ],
  ]);

  return {
    content: [
      {
        type: "text",
        text: `Email marked as ${isRead ? "read" : "unread"}`,
      },
    ],
  };
}

// Register all tools on a fresh server instance
function registerTools(server) {
  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: "list_inbox",
        description: "List messages in the inbox with subject, from, date, and preview",
        inputSchema: {
          type: "object",
          properties: {
            limit: {
              type: "number",
              description: "Maximum number of messages to return (default: 20)",
              default: 20,
            },
          },
        },
      },
      {
        name: "search_email",
        description: "Search messages by keyword, sender, or subject",
        inputSchema: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "Search query text",
            },
            sender: {
              type: "string",
              description: "Filter by sender email address",
            },
            subject: {
              type: "string",
              description: "Filter by subject text",
            },
            limit: {
              type: "number",
              description: "Maximum number of results (default: 20)",
              default: 20,
            },
          },
        },
      },
      {
        name: "read_email",
        description: "Read the full content of an email by its ID",
        inputSchema: {
          type: "object",
          properties: {
            emailId: {
              type: "string",
              description: "The email ID to read",
            },
          },
          required: ["emailId"],
        },
      },
      {
        name: "send_email",
        description: "Send an email",
        inputSchema: {
          type: "object",
          properties: {
            to: {
              type: "array",
              items: { type: "string" },
              description: "Recipient email addresses",
            },
            subject: {
              type: "string",
              description: "Email subject",
            },
            body: {
              type: "string",
              description: "Email body (plain text or HTML)",
            },
            fromAlias: {
              type: "string",
              description: "From identity/alias ID (optional)",
            },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "list_aliases",
        description: "List all Fastmail identities/aliases",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "list_folders",
        description: "List all mailboxes/folders",
        inputSchema: {
          type: "object",
          properties: {},
        },
      },
      {
        name: "move_email",
        description: "Move an email to a different folder",
        inputSchema: {
          type: "object",
          properties: {
            emailId: {
              type: "string",
              description: "The email ID to move",
            },
            mailboxId: {
              type: "string",
              description: "The destination mailbox/folder ID",
            },
          },
          required: ["emailId", "mailboxId"],
        },
      },
      {
        name: "delete_email",
        description: "Move an email to trash",
        inputSchema: {
          type: "object",
          properties: {
            emailId: {
              type: "string",
              description: "The email ID to delete",
            },
          },
          required: ["emailId"],
        },
      },
      {
        name: "mark_read",
        description: "Mark an email as read or unread",
        inputSchema: {
          type: "object",
          properties: {
            emailId: {
              type: "string",
              description: "The email ID",
            },
            isRead: {
              type: "boolean",
              description: "True to mark as read, false for unread",
            },
          },
          required: ["emailId", "isRead"],
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      const accountId = await getAccountId();

      switch (name) {
        case "list_inbox":
          return await listInbox(accountId, args.limit || 20);
        case "search_email":
          return await searchEmail(accountId, args);
        case "read_email":
          return await readEmail(accountId, args.emailId);
        case "send_email":
          return await sendEmail(accountId, args);
        case "list_aliases":
          return await listAliases(accountId);
        case "list_folders":
          return await listFolders(accountId);
        case "move_email":
          return await moveEmail(accountId, args.emailId, args.mailboxId);
        case "delete_email":
          return await deleteEmail(accountId, args.emailId);
        case "mark_read":
          return await markRead(accountId, args.emailId, args.isRead);
        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error: ${error.message}`,
          },
        ],
      };
    }
  });
}

// Create HTTP server
const httpServer = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  
  if (parsedUrl.pathname === '/mcp') {
    console.log(`[${new Date().toISOString()}] ${req.method} /mcp from ${req.socket.remoteAddress}`);
    
    if (req.method === 'GET') {
      // Handle non-SSE GET requests (e.g., health probes) gracefully
      const acceptHeader = req.headers['accept'] || '';
      if (!acceptHeader.includes('text/event-stream')) {
        // Non-SSE probe - return 200 OK
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ status: 'ok', server: 'fastmail-mcp' }));
        return;
      }
      
      // SSE request - return 405 (not supported in this version)
      res.writeHead(405, { 
        'Content-Type': 'application/json',
        'Allow': 'POST'
      });
      res.end(JSON.stringify({
        jsonrpc: '2.0',
        error: {
          code: -32000,
          message: 'Method not allowed. Use POST for MCP requests.'
        },
        id: null
      }));
      return;
    }
    
    if (req.method === 'POST') {
      try {
        // Parse request body
        let body = '';
        req.on('data', chunk => {
          body += chunk.toString();
        });
        
        await new Promise((resolve) => {
          req.on('end', resolve);
        });
        
        const parsedBody = body ? JSON.parse(body) : undefined;
        
        // Create a fresh server instance for this request
        const server = new Server(
          {
            name: "fastmail-mcp",
            version: "1.0.0",
          },
          {
            capabilities: {
              tools: {},
            },
          }
        );
        
        // Register all tools on this fresh instance
        registerTools(server);
        
        // Create a fresh stateless transport for this request
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: undefined // Stateless mode
        });
        
        await server.connect(transport);
        await transport.handleRequest(req, res, parsedBody);
        
        // Clean up transport after request
        res.on('close', () => {
          transport.close();
        });
      } catch (error) {
        console.error(`[${new Date().toISOString()}] Error handling /mcp request:`, error);
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ 
            jsonrpc: '2.0',
            error: {
              code: -32603,
              message: 'Internal server error'
            },
            id: null
          }));
        }
      }
    } else {
      res.writeHead(405, { 
        'Content-Type': 'application/json',
        'Allow': 'POST'
      });
      res.end(JSON.stringify({
        jsonrpc: '2.0',
        error: {
          code: -32000,
          message: `Method ${req.method} not allowed. Use POST for MCP requests.`
        },
        id: null
      }));
    }
  } else if (parsedUrl.pathname === '/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', service: 'fastmail-mcp' }));
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found. Available endpoints: /mcp (POST), /health');
  }
});

httpServer.listen(PORT, () => {
  console.log(`Fastmail MCP Server running on http://localhost:${PORT}`);
  console.log(`Streamable HTTP endpoint at http://localhost:${PORT}/mcp`);
  console.log(`Health check at http://localhost:${PORT}/health`);
});

process.on('SIGINT', async () => {
  console.log('\nShutting down server...');
  httpServer.close();
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('\nShutting down server...');
  httpServer.close();
  process.exit(0);
});
