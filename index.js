#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
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

// CalDAV credentials — mirror sent .ics invites onto Ryan's primary calendar.
// Defaults are Ryan's setup; override via env when deploying for others.
const CALDAV_USER = process.env.FASTMAIL_CALDAV_USER || "ryan@symbio.live";
const CALDAV_PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD || "7f9q4x679y555n7g";
const CALDAV_BASE = "https://caldav.fastmail.com";
const CALDAV_CALENDAR_PATH = process.env.FASTMAIL_CALDAV_CALENDAR_PATH || "/dav/calendars/user/ryan@symbio.live/F7F39F26-41B5-11F1-880A-F1376648A29D/";

async function putCalendarEvent(icsContent) {
  if (!CALDAV_USER || !CALDAV_PASSWORD) return { ok: false, skipped: true };
  const uidMatch = icsContent.match(/^UID:(.+)$/m);
  if (!uidMatch) return { ok: false, error: "no UID in ics" };
  const uid = uidMatch[1].trim();
  // Strip METHOD line — METHOD is for transport (iTIP), not for stored events
  const eventIcs = icsContent.replace(/^METHOD:.*\r?\n/m, "");
  const auth = Buffer.from(`${CALDAV_USER}:${CALDAV_PASSWORD}`).toString("base64");
  const safeUid = encodeURIComponent(uid);
  const url = `${CALDAV_BASE}${CALDAV_CALENDAR_PATH}${safeUid}.ics`;
  try {
    const res = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Basic ${auth}`,
        "Content-Type": "text/calendar; charset=utf-8",
      },
      body: eventIcs,
    });
    if (!res.ok && res.status !== 412 && res.status !== 204) {
      const text = await res.text();
      console.error("[CalDAV] PUT failed:", res.status, text.slice(0, 200));
      return { ok: false, status: res.status };
    }
    return { ok: true, uid, status: res.status };
  } catch (e) {
    console.error("[CalDAV] PUT error:", e.message);
    return { ok: false, error: e.message };
  }
}

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
async function listInbox(accountId, limit, folderId) {
  let mailboxId = folderId;

  // If no folderId given, default to the inbox (role === "inbox")
  if (!mailboxId) {
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
    mailboxId = inboxMailbox.id;
  }

  // Then query emails in the target mailbox
  const response = await jmapRequest([
    [
      "Email/query",
      {
        accountId,
        filter: { inMailbox: mailboxId },
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

const BRAND_COLORS = {
  'injury.media': '#c8a45a',
  'transcendmedia.agency': '#2563eb',
};

const SIGNOFF_RE = /^(talk soon|best|cheers|warm regards|regards|sincerely|thanks|thank you|looking forward)[,!.]?\s*$/i;

function formatEmailBody(body, fromEmail, ctaUrl, ctaLabel) {
  const domain = (fromEmail || '').split('@')[1] || '';
  const color = BRAND_COLORS[domain] || '#2563eb';

  const outerOpen = `<!DOCTYPE html><html><body style="margin:0;padding:0;background-color:#f6f6f6;">
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f6f6f6;padding:24px 0;">
  <tr><td align="center">
    <table cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background-color:#ffffff;border-radius:8px;overflow:hidden;">
      <tr><td style="padding:24px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:15px;line-height:1.6;color:#333333;">`;

  const outerClose = `      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>`;

  const glassCard = (content) =>
    `<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:0 0 16px 0;">
  <tr>
    <td style="background-color:rgba(245,245,247,0.6);border:1px solid rgba(0,0,0,0.06);border-radius:12px;padding:16px 20px;">
      <p style="margin:0;font-size:15px;line-height:1.6;color:#1a1a1a;">${content}</p>
    </td>
  </tr>
</table>`;

  const ctaCard = (url, label) =>
    `<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:0 0 16px 0;">
  <tr>
    <td align="center" style="background-color:${color}18;border:1px solid ${color}40;border-radius:14px;padding:20px 24px;max-width:480px;">
      <a href="${url}" style="display:inline-block;background-color:${color};color:#ffffff;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:15px;font-weight:600;text-decoration:none;padding:10px 24px;border-radius:8px;">${label || 'View Now'} →</a>
    </td>
  </tr>
</table>`;

  const chunks = body.split(/\n\n+/);
  let html = '';
  let ctaInserted = false;

  for (const chunk of chunks) {
    const trimmed = chunk.trim();
    if (!trimmed) continue;

    if (trimmed.includes('{{CTA}}')) {
      if (ctaUrl) {
        html += ctaCard(ctaUrl, ctaLabel);
        ctaInserted = true;
      }
      continue;
    }

    if (SIGNOFF_RE.test(trimmed)) {
      html += `<p style="margin:0 0 12px 0;font-size:15px;line-height:1.6;color:#333333;">${trimmed}</p>`;
      continue;
    }

    // Convert single newlines within a chunk to <br>
    const inner = trimmed.replace(/\n/g, '<br>');
    html += glassCard(inner);
  }

  // Append CTA after last paragraph if {{CTA}} wasn't in the body
  if (ctaUrl && !ctaInserted) {
    html += ctaCard(ctaUrl, ctaLabel);
  }

  return { outerOpen, outerClose, bodyHtml: html, color };
}

async function sendEmail(accountId, args) {
  const emailId = `draft-${Date.now()}`;
  
  // Get the drafts mailbox
  const mailboxResponse = await jmapRequest([
    [
      "Mailbox/get",
      {
        accountId,
      },
      "0",
    ],
  ]);

  const draftsMailbox = mailboxResponse.methodResponses[0][1].list.find(
    (mb) => mb.role === "drafts"
  );

  if (!draftsMailbox) {
    throw new Error("Drafts mailbox not found");
  }

  // Get all identities to search by email or ID
  const identityResponse = await jmapRequest([["Identity/get", { accountId }, "0"]]);
  const identities = identityResponse.methodResponses[0][1].list;
  
  if (!identities || identities.length === 0) {
    throw new Error("No identities found");
  }

  // Find the identity: if fromAlias looks like an email, match by email field; otherwise match by ID
  let selectedIdentity;
  if (args.fromAlias) {
    if (args.fromAlias.includes('@')) {
      // fromAlias is an email address, search by email
      selectedIdentity = identities.find(id => id.email === args.fromAlias);
      if (!selectedIdentity) {
        throw new Error(`Identity with email ${args.fromAlias} not found`);
      }
    } else {
      // fromAlias is an ID, search by id
      selectedIdentity = identities.find(id => id.id === args.fromAlias);
      if (!selectedIdentity) {
        throw new Error(`Identity with ID ${args.fromAlias} not found`);
      }
    }
  } else {
    // No fromAlias specified, use the first identity
    selectedIdentity = identities[0];
  }

  const fromEmail = [{ email: selectedIdentity.email }];
  const identityId = selectedIdentity.id;

  // Build email body with glassmorphic card formatting
  const { outerOpen, outerClose, bodyHtml } = formatEmailBody(
    args.body,
    selectedIdentity.email,
    args.ctaUrl,
    args.ctaLabel
  );

  let signatureBlock = '';
  if (selectedIdentity.htmlSignature) {
    signatureBlock = `<hr style="border:none;border-top:1px solid #e5e5e5;margin:24px 0;">${selectedIdentity.htmlSignature}`;
  }

  let emailBody = `${outerOpen}${bodyHtml}${signatureBlock}${outerClose}`;

  // Build email structure with or without attachments
  let emailStructure;
  const bodyValues = {
    body: { value: emailBody, charset: "utf-8", isTruncated: false },
  };

  if (args.attachments && args.attachments.length > 0) {
    // Separate .ics calendar invites from regular attachments
    const calendarInvites = [];
    const regularAttachments = [];
    
    args.attachments.forEach((attachment, index) => {
      if (attachment.filename && attachment.filename.toLowerCase().endsWith('.ics')) {
        calendarInvites.push({ ...attachment, index });
      } else {
        regularAttachments.push({ ...attachment, index });
      }
    });

    // Handle calendar invites as inline multipart/alternative
    if (calendarInvites.length > 0) {
      const calendarInvite = calendarInvites[0]; // Take first .ics
      const calPartId = `calendar${calendarInvite.index}`;
      
      // Decode base64 to get calendar data
      const calendarData = Buffer.from(calendarInvite.data, 'base64').toString('utf-8');
      
      // Add calendar data to bodyValues
      bodyValues[calPartId] = {
        value: calendarData,
        charset: "utf-8",
        isTruncated: false,
      };

      // Create plain text fallback
      const textFallback = `Calendar Invite\n\n${args.subject}\n\nThis is a calendar invitation. Please use a calendar-enabled email client to view and respond.`;
      bodyValues.textFallback = {
        value: textFallback,
        charset: "utf-8",
        isTruncated: false,
      };

      // Body alternative (text/plain + text/html only — calendar is NOT
      // a body alternative; it's a sibling part inside multipart/mixed so
      // Apple Mail recognises it as an actionable invite, not a rendering).
      const bodyAlternative = {
        type: "multipart/alternative",
        subParts: [
          { partId: "textFallback", type: "text/plain" },
          { partId: "body", type: "text/html" },
        ],
      };

      // Calendar part: inline (not attachment) so Apple Mail recognises
      // it as actionable. METHOD:REQUEST is set inside the .ics body —
      // Fastmail's JMAP doesn't accept `charset` or `headers` on body
      // sub-parts, so we can't override the wire-level Content-Type
      // parameters. Apple Mail falls back to reading METHOD from the
      // .ics body, which is sufficient for Accept/Decline rendering.
      const calendarPart = {
        partId: calPartId,
        type: "text/calendar",
        disposition: "inline",
      };

      // Optional regular attachments alongside the invite
      const attachmentParts = regularAttachments.map((attachment) => {
        const partId = `attachment${attachment.index}`;

        bodyValues[partId] = {
          value: attachment.data,
          charset: null,
          isTruncated: false,
        };

        const cleanType = (attachment.type || "application/octet-stream").split(';')[0].trim();

        return {
          partId: partId,
          type: cleanType,
          name: attachment.filename,
          disposition: "attachment",
          cid: null,
        };
      });

      emailStructure = {
        mailboxIds: { [draftsMailbox.id]: true },
        from: fromEmail,
        to: args.to.map((email) => ({ email })),
        subject: args.subject,
        bodyStructure: {
          type: "multipart/mixed",
          subParts: [
            bodyAlternative,
            calendarPart,
            ...attachmentParts,
          ],
        },
        bodyValues: bodyValues,
      };
    } else {
      // No calendar invites, only regular attachments
      const attachmentParts = regularAttachments.map((attachment) => {
        const partId = `attachment${attachment.index}`;
        
        bodyValues[partId] = {
          value: attachment.data,
          charset: null,
          isTruncated: false,
        };

        const cleanType = (attachment.type || "application/octet-stream").split(';')[0].trim();

        return {
          partId: partId,
          type: cleanType,
          name: attachment.filename,
          disposition: "attachment",
          cid: null,
        };
      });

      emailStructure = {
        mailboxIds: { [draftsMailbox.id]: true },
        from: fromEmail,
        to: args.to.map((email) => ({ email })),
        subject: args.subject,
        bodyStructure: {
          type: "multipart/mixed",
          subParts: [
            {
              partId: "body",
              type: "text/html",
            },
            ...attachmentParts,
          ],
        },
        bodyValues: bodyValues,
      };
    }
  } else {
    // Without attachments: simple HTML email
    emailStructure = {
      mailboxIds: { [draftsMailbox.id]: true },
      from: fromEmail,
      to: args.to.map((email) => ({ email })),
      subject: args.subject,
      htmlBody: [{ partId: "body", type: "text/html" }],
      bodyValues: bodyValues,
    };
  }

  // Create email in drafts mailbox, then submit it
  const response = await jmapRequest([
    [
      "Email/set",
      {
        accountId,
        create: {
          [emailId]: emailStructure,
        },
      },
      "0",
    ],
    [
      "EmailSubmission/set",
      {
        accountId,
        onSuccessDestroyEmail: ["#submission1"],
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

  // Check if Email/set succeeded
  const emailSetResult = response.methodResponses[0][1];
  if (emailSetResult.notCreated?.[emailId]) {
    const error = emailSetResult.notCreated[emailId];
    console.error('[sendEmail] Email/set failed. Full response:', JSON.stringify(response.methodResponses, null, 2));
    throw new Error(`Email creation failed: ${error.description || JSON.stringify(error)}`);
  }

  // Check for submission errors
  const submissionResult = response.methodResponses[1][1];
  if (submissionResult.notCreated) {
    const error = submissionResult.notCreated.submission1;
    console.error('[sendEmail] EmailSubmission/set failed. Full response:', JSON.stringify(response.methodResponses, null, 2));
    throw new Error(`Email submission failed: ${error.description || JSON.stringify(error)}`);
  }

  const attachmentCount = args.attachments ? args.attachments.length : 0;
  let message = attachmentCount > 0
    ? `Email sent successfully with ${attachmentCount} attachment${attachmentCount > 1 ? 's' : ''}`
    : "Email sent successfully";

  // Mirror any .ics invite onto the organizer's CalDAV calendar so the
  // event appears on their phone calendar app (Fastmail JMAP doesn't
  // expose calendar APIs — must be done via CalDAV).
  if (args.attachments?.length) {
    for (const att of args.attachments) {
      if (att.filename?.toLowerCase().endsWith('.ics') && att.data) {
        try {
          const ics = Buffer.from(att.data, 'base64').toString('utf-8');
          const result = await putCalendarEvent(ics);
          if (result.ok) message += " (added to calendar)";
          else if (!result.skipped) console.error('[sendEmail] CalDAV mirror failed:', result);
        } catch (e) {
          console.error('[sendEmail] CalDAV mirror exception:', e.message);
        }
      }
    }
  }

  return {
    content: [
      {
        type: "text",
        text: message,
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
        description: "List messages from a Fastmail folder (defaults to the inbox if no folderId is provided). Returns subject, from, date, and preview. Use list_folders first to discover folder IDs for non-inbox folders.",
        inputSchema: {
          type: "object",
          properties: {
            limit: {
              type: "number",
              description: "Maximum number of messages to return (default: 20)",
              default: 20,
            },
            folderId: {
              type: "string",
              description: "Optional Fastmail folder/mailbox ID. Omit to list the inbox.",
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
        description: "Send an email with optional file attachments (supports .ics calendar invites)",
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
              description: "From identity/alias ID or email address (optional)",
            },
            ctaUrl: {
              type: "string",
              description: "Optional CTA button URL. Inserts a branded call-to-action card at {{CTA}} in the body, or appended after the last paragraph.",
            },
            ctaLabel: {
              type: "string",
              description: "Label for the CTA button (default: 'View Now')",
            },
            attachments: {
              type: "array",
              description: "Optional array of file attachments. For .ics files, automatically sets Content-Type to 'text/calendar; method=REQUEST' to trigger 'Add to Calendar' in mail clients.",
              items: {
                type: "object",
                properties: {
                  filename: {
                    type: "string",
                    description: "Name of the file (e.g., 'invite.ics')",
                  },
                  type: {
                    type: "string",
                    description: "MIME type (e.g., 'application/pdf', 'image/png'). For .ics files, this is automatically set to 'text/calendar; method=REQUEST'.",
                  },
                  data: {
                    type: "string",
                    description: "Base64-encoded file data",
                  },
                },
                required: ["filename", "data"],
              },
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
          return await listInbox(accountId, args.limit || 20, args.folderId);
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

// Stdio mode: spawned by Claude Code (or any MCP client) on demand
if (process.argv.includes('--stdio')) {
  const server = new Server(
    { name: "fastmail-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );
  registerTools(server);
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Fastmail MCP server running on stdio");
} else {

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

} // end http branch
