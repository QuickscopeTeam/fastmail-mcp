#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import fetch from "node-fetch";
import http from "http";
import url from "url";

const FASTMAIL_API_URL = "https://api.fastmail.com/jmap/api/";
const API_KEY = process.env.FASTMAIL_API_KEY || "fmu1-d01e43a8-f3ca5b7579eb5aac1f2df46b23440060-0-e12f16ead889af9da7e5dbf2720a49ff";
const PORT = process.env.PORT || 3000;

class FastmailMCPServer {
  constructor() {
    this.server = new Server(
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

    this.accountId = null;
    this.setupHandlers();
    this.server.onerror = (error) => console.error("[MCP Error]", error);
  }

  async jmapRequest(methodCalls) {
    const response = await fetch(FASTMAIL_API_URL, {
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

  async getAccountId() {
    if (this.accountId) return this.accountId;

    const response = await this.jmapRequest([
      ["Session/get", {}, "0"],
    ]);

    const sessionData = response.methodResponses[0][1];
    const primaryAccounts = sessionData.primaryAccounts;
    this.accountId = primaryAccounts["urn:ietf:params:jmap:mail"];
    return this.accountId;
  }

  setupHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
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

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        const accountId = await this.getAccountId();

        switch (name) {
          case "list_inbox":
            return await this.listInbox(accountId, args.limit || 20);
          case "search_email":
            return await this.searchEmail(accountId, args);
          case "read_email":
            return await this.readEmail(accountId, args.emailId);
          case "send_email":
            return await this.sendEmail(accountId, args);
          case "list_aliases":
            return await this.listAliases(accountId);
          case "list_folders":
            return await this.listFolders(accountId);
          case "move_email":
            return await this.moveEmail(accountId, args.emailId, args.mailboxId);
          case "delete_email":
            return await this.deleteEmail(accountId, args.emailId);
          case "mark_read":
            return await this.markRead(accountId, args.emailId, args.isRead);
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

  async listInbox(accountId, limit) {
    const response = await this.jmapRequest([
      [
        "Email/query",
        {
          accountId,
          filter: { inMailbox: null },
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

  async searchEmail(accountId, args) {
    const filter = {};
    if (args.query) filter.text = args.query;
    if (args.sender) filter.from = args.sender;
    if (args.subject) filter.subject = args.subject;

    const response = await this.jmapRequest([
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

  async readEmail(accountId, emailId) {
    const response = await this.jmapRequest([
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

  async sendEmail(accountId, args) {
    const emailId = `draft-${Date.now()}`;
    const identityId = args.fromAlias || null;

    const response = await this.jmapRequest([
      [
        "Email/set",
        {
          accountId,
          create: {
            [emailId]: {
              mailboxIds: {},
              from: identityId ? [{ email: identityId }] : undefined,
              to: args.to.map((email) => ({ email })),
              subject: args.subject,
              textBody: [{ partId: "body", type: "text/plain" }],
              bodyValues: {
                body: { value: args.body },
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

    return {
      content: [
        {
          type: "text",
          text: "Email sent successfully",
        },
      ],
    };
  }

  async listAliases(accountId) {
    const response = await this.jmapRequest([
      [
        "Identity/get",
        {
          accountId,
        },
        "0",
      ],
    ]);

    const identities = response.methodResponses[0][1].list;
    const formatted = identities.map((identity) => ({
      id: identity.id,
      name: identity.name,
      email: identity.email,
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

  async listFolders(accountId) {
    const response = await this.jmapRequest([
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

  async moveEmail(accountId, emailId, mailboxId) {
    const response = await this.jmapRequest([
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

  async deleteEmail(accountId, emailId) {
    const mailboxResponse = await this.jmapRequest([
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

    return await this.moveEmail(accountId, emailId, trashMailbox.id);
  }

  async markRead(accountId, emailId, isRead) {
    const response = await this.jmapRequest([
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

  async run() {
    const httpServer = http.createServer(async (req, res) => {
      const parsedUrl = url.parse(req.url, true);
      
      if (parsedUrl.pathname === '/mcp') {
        // Handle Streamable HTTP (POST requests with JSON-RPC)
        console.log(`[${new Date().toISOString()}] ${req.method} /mcp from ${req.socket.remoteAddress}`);
        
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
          
          // Create a fresh stateless transport for each request
          const mcpTransport = new StreamableHTTPServerTransport({
            sessionIdGenerator: undefined // Stateless mode
          });
          await this.server.connect(mcpTransport);
          
          await mcpTransport.handleRequest(req, res, parsedBody);
          
          // Clean up transport after request
          res.on('close', () => {
            mcpTransport.close();
          });
        } catch (error) {
          console.error(`[${new Date().toISOString()}] Error handling /mcp request:`, error);
          if (!res.headersSent) {
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Internal server error' }));
          }
        }
      } else if (parsedUrl.pathname === '/sse') {
        // Handle SSE (Server-Sent Events) for legacy clients
        console.log(`[${new Date().toISOString()}] New SSE connection from ${req.socket.remoteAddress}`);
        
        const transport = new SSEServerTransport(parsedUrl.pathname, res);
        await this.server.connect(transport);
        
        req.on('close', () => {
          console.log(`[${new Date().toISOString()}] SSE connection closed`);
        });
      } else if (parsedUrl.pathname === '/health') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ status: 'ok', service: 'fastmail-mcp' }));
      } else {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Not Found. Available endpoints: /mcp (Streamable HTTP), /sse (SSE), /health');
      }
    });

    httpServer.listen(PORT, () => {
      console.log(`Fastmail MCP Server running on http://localhost:${PORT}`);
      console.log(`Streamable HTTP endpoint at http://localhost:${PORT}/mcp`);
      console.log(`SSE endpoint available at http://localhost:${PORT}/sse`);
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
  }
}

const server = new FastmailMCPServer();
server.run().catch(console.error);
