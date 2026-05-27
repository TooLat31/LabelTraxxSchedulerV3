import { createClient } from "@supabase/supabase-js";
import { Resend } from "resend";

const ATTACHMENT_BUCKET = "labeltraxx-attachments";
const SHARED_STATE_ROW_ID = "labeltraxx-shared-state";
const ACTIVITY_LOG_LIMIT = 300;

export const config = {
  api: {
    bodyParser: false,
  },
};

function safeText(value) {
  if (value == null) return "";
  return String(value).trim();
}

function comparable(value) {
  return safeText(value).toLowerCase();
}

function makeId(prefix) {
  if (globalThis.crypto?.randomUUID) {
    return `${prefix}-${globalThis.crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function sanitizeAttachmentName(name) {
  const cleaned = safeText(name).replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/-+/g, "-");
  return cleaned.replace(/^-|-$/g, "") || "file";
}

function parseAddressParts(rawValue) {
  const raw = safeText(rawValue);
  const match = raw.match(/^(.*?)(?:<([^>]+)>)?$/);
  const name = safeText(match?.[1]).replace(/^"|"$/g, "");
  const email = safeText(match?.[2] || raw).replace(/^<|>$/g, "");
  return { name, email };
}

function extractJobNumber(...values) {
  for (const value of values) {
    const match = safeText(value).match(/\b\d{4,8}\b/);
    if (match) return match[0];
  }
  return "";
}

function stripHtml(html) {
  const raw = safeText(html);
  if (!raw) return "";
  return raw
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeEmailBody(text) {
  return safeText(text)
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.replace(/\s+$/g, ""))
    .filter((line) => !/^\[cid:.*\]$/i.test(line))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function parseForwardedHeaders(lines, startIndex) {
  const headers = {};
  let index = startIndex;
  let activeHeader = "";

  while (index < lines.length) {
    const line = lines[index];
    if (!safeText(line)) {
      index += 1;
      break;
    }

    const match = line.match(/^([A-Za-z-]+):\s*(.*)$/);
    if (match) {
      activeHeader = match[1];
      headers[activeHeader] = safeText(match[2]);
    } else if (activeHeader) {
      headers[activeHeader] = `${headers[activeHeader]} ${safeText(line)}`.trim();
    } else {
      break;
    }
    index += 1;
  }

  return {
    headers,
    bodyStartIndex: index,
  };
}

function extractOriginalMessage(email) {
  const rawBody = normalizeEmailBody(safeText(email.text) || stripHtml(email.html));
  if (!rawBody) {
    return {
      senderRaw: safeText(email.from),
      subject: safeText(email.subject),
      to: Array.isArray(email.to) ? email.to : [],
      cc: Array.isArray(email.cc) ? email.cc : [],
      bodyText: "",
      wasForwarded: false,
    };
  }

  const lines = rawBody.split("\n");
  const fromIndexes = [];
  lines.forEach((line, index) => {
    if (/^From:\s/i.test(line)) fromIndexes.push(index);
  });

  for (let index = fromIndexes.length - 1; index >= 0; index -= 1) {
    const fromIndex = fromIndexes[index];
    const { headers, bodyStartIndex } = parseForwardedHeaders(lines, fromIndex);
    if (!safeText(headers.From) || !safeText(headers.Subject)) continue;

    const originalBody = lines
      .slice(bodyStartIndex)
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    return {
      senderRaw: safeText(headers.From),
      subject: safeText(headers.Subject),
      to: safeText(headers.To)
        ? safeText(headers.To)
            .split(/;\s*|,\s*/)
            .map((value) => safeText(value))
            .filter(Boolean)
        : [],
      cc: safeText(headers.Cc)
        ? safeText(headers.Cc)
            .split(/;\s*|,\s*/)
            .map((value) => safeText(value))
            .filter(Boolean)
        : [],
      bodyText: originalBody || rawBody,
      wasForwarded: true,
    };
  }

  return {
    senderRaw: safeText(email.from),
    subject: safeText(email.subject),
    to: Array.isArray(email.to) ? email.to : [],
    cc: Array.isArray(email.cc) ? email.cc : [],
    bodyText: rawBody,
    wasForwarded: false,
  };
}

function appendActivityEntry(currentLog, entry) {
  const nextEntry = {
    id: entry.id || makeId("activity"),
    actor: safeText(entry.actor) || "system",
    action: safeText(entry.action),
    scope: safeText(entry.scope) || "General",
    description: safeText(entry.description),
    createdAt: entry.createdAt || new Date().toISOString(),
    details: entry.details && typeof entry.details === "object" ? entry.details : {},
  };
  return [nextEntry, ...(Array.isArray(currentLog) ? currentLog : [])].slice(0, ACTIVITY_LOG_LIMIT);
}

async function readRawBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

function buildSupabaseClient() {
  const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
  const supabaseKey =
    process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.VITE_SUPABASE_ANON_KEY || process.env.SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseKey) {
    throw new Error("Missing Supabase server configuration.");
  }

  return createClient(supabaseUrl, supabaseKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  });
}

function buildResendClient() {
  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) throw new Error("Missing RESEND_API_KEY.");
  return new Resend(apiKey);
}

function inferAssignedUser(users, recipients) {
  const knownUsers = Array.isArray(users) ? users : [];
  for (const recipient of recipients || []) {
    const email = parseAddressParts(recipient).email;
    const localPart = safeText(email.split("@")[0]).split("+")[0];
    if (!localPart) continue;
    const match = knownUsers.find((user) => comparable(user.username) === comparable(localPart));
    if (match) return match.username;
  }
  return "";
}

function isInlineAttachment(attachment) {
  const filename = safeText(attachment.filename || attachment.name).toLowerCase();
  const disposition = safeText(attachment.content_disposition).toLowerCase();
  const contentId = safeText(attachment.content_id);
  return disposition === "inline" || !!contentId || /^image\d+\.(png|jpe?g|gif|webp)$/i.test(filename);
}

function buildRequestDescription(emailView, attachmentSummaries) {
  const lines = [
    `Subject: ${safeText(emailView.subject) || "(no subject)"}`,
    `From: ${safeText(emailView.senderRaw) || "-"}`,
  ];

  if (Array.isArray(emailView.cc) && emailView.cc.length) {
    lines.push(`CC: ${emailView.cc.map((value) => safeText(value)).filter(Boolean).join(", ")}`);
  }

  if (attachmentSummaries.length) {
    lines.push(`Attachments: ${attachmentSummaries.map((item) => item.name).join(", ")}`);
  }

  if (emailView.bodyText) {
    lines.push("", emailView.bodyText);
  }

  return lines.join("\n").trim();
}

async function uploadInboundAttachments({ supabase, resend, emailId, senderTag }) {
  const { data: attachmentList, error } = await resend.emails.receiving.attachments.list({
    emailId,
  });

  if (error) {
    console.warn("Failed to list inbound attachments.", error.message || error);
    return [];
  }

  const attachments = Array.isArray(attachmentList?.data) ? attachmentList.data : [];
  const uploaded = [];

  for (const attachment of attachments) {
    if (isInlineAttachment(attachment)) continue;
    const downloadUrl = safeText(attachment.download_url);
    if (!downloadUrl) continue;

    try {
      const response = await fetch(downloadUrl);
      if (!response.ok) {
        throw new Error(`Failed to download inbound attachment ${safeText(attachment.filename)}.`);
      }

      const buffer = Buffer.from(await response.arrayBuffer());
      const storagePath = `inbound-email/${safeText(emailId)}/${makeId("att")}-${sanitizeAttachmentName(attachment.filename)}`;
      const contentType = safeText(attachment.content_type) || response.headers.get("content-type") || "application/octet-stream";

      const { error: uploadError } = await supabase.storage.from(ATTACHMENT_BUCKET).upload(storagePath, buffer, {
        upsert: false,
        contentType,
      });

      if (uploadError) {
        throw new Error(uploadError.message || `Failed to upload ${safeText(attachment.filename)} to storage.`);
      }

      const { data } = supabase.storage.from(ATTACHMENT_BUCKET).getPublicUrl(storagePath);
      uploaded.push({
        id: makeId("att"),
        name: safeText(attachment.filename) || "attachment",
        size: Number(attachment.size) || buffer.length,
        type: contentType,
        uploadedAt: new Date().toISOString(),
        uploadedBy: senderTag,
        storagePath,
        publicUrl: safeText(data?.publicUrl),
        dataUrl: "",
      });
    } catch (attachmentError) {
      console.warn(`Falling back to Resend attachment link for ${safeText(attachment.filename)}.`, attachmentError);
      uploaded.push({
        id: makeId("att"),
        name: safeText(attachment.filename) || "attachment",
        size: Number(attachment.size) || 0,
        type: safeText(attachment.content_type) || "application/octet-stream",
        uploadedAt: new Date().toISOString(),
        uploadedBy: senderTag,
        storagePath: "",
        publicUrl: downloadUrl,
        dataUrl: "",
      });
    }
  }

  return uploaded;
}

async function getSharedSnapshot(supabase) {
  const { data, error } = await supabase
    .from("app_state")
    .select("payload")
    .eq("id", SHARED_STATE_ROW_ID)
    .maybeSingle();

  if (error) throw new Error(error.message || "Failed to load shared scheduler state.");
  return data?.payload && typeof data.payload === "object" ? data.payload : {};
}

async function saveSharedSnapshot(supabase, payload, updatedBy) {
  const { error } = await supabase.from("app_state").upsert({
    id: SHARED_STATE_ROW_ID,
    payload,
    updated_by: safeText(updatedBy) || "email-bot",
  });

  if (error) throw new Error(error.message || "Failed to save shared scheduler state.");
}

export default async function handler(req, res) {
  if (req.method === "GET") {
    res.status(200).json({
      ok: true,
      route: "resend-inbound",
      date: new Date().toISOString(),
    });
    return;
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    res.status(405).json({ error: "Method not allowed." });
    return;
  }

  try {
    const resend = buildResendClient();
    const supabase = buildSupabaseClient();
    const rawBody = await readRawBody(req);
    const webhookSecret = process.env.RESEND_WEBHOOK_SECRET;

    let event;
    if (webhookSecret) {
      event = resend.webhooks.verify({
        payload: rawBody,
        headers: {
          id: safeText(req.headers["svix-id"]),
          timestamp: safeText(req.headers["svix-timestamp"]),
          signature: safeText(req.headers["svix-signature"]),
        },
        webhookSecret,
      });
    } else {
      event = JSON.parse(rawBody || "{}");
    }

    if (event?.type !== "email.received") {
      res.status(200).json({ ok: true, ignored: true, type: safeText(event?.type) });
      return;
    }

    const emailId = safeText(event.data?.email_id);
    if (!emailId) {
      res.status(400).json({ error: "Missing email_id on webhook payload." });
      return;
    }

    const currentSnapshot = await getSharedSnapshot(supabase);
    const existingRequests = Array.isArray(currentSnapshot.requests) ? currentSnapshot.requests : [];

    if (existingRequests.some((request) => comparable(request.inboundEmailId) === comparable(emailId))) {
      res.status(200).json({ ok: true, duplicate: true, emailId });
      return;
    }

    const { data: email, error: emailError } = await resend.emails.receiving.get(emailId);
    if (emailError || !email) {
      throw new Error(emailError?.message || "Failed to retrieve inbound email content.");
    }

    const recipients = Array.isArray(email.to) ? email.to : [];
    const sender = parseAddressParts(email.from);
    const senderLabel = sender.name || sender.email || "Inbound email";
    const emailView = extractOriginalMessage(email);
    const originalSender = parseAddressParts(emailView.senderRaw);
    const originalSenderLabel = originalSender.name || originalSender.email || senderLabel;
    const attachmentUploads = await uploadInboundAttachments({
      supabase,
      resend,
      emailId,
      senderTag: originalSenderLabel,
    });

    const users = Array.isArray(currentSnapshot.users) ? currentSnapshot.users : [];
    const assignedToAccount = inferAssignedUser(users, recipients);
    const jobNumber =
      extractJobNumber(emailView.subject, emailView.bodyText, email.subject) ||
      `EMAIL-${safeText(emailId).replace(/[^a-zA-Z0-9]/g, "").slice(0, 6).toUpperCase()}`;

    const newRequest = {
      id: makeId("req"),
      jobNumber,
      customer: originalSender.name || originalSender.email || sender.name || sender.email || "Email Request",
      requestorName: originalSender.email ? `${originalSenderLabel} <${originalSender.email}>` : originalSenderLabel,
      description: buildRequestDescription(emailView, attachmentUploads),
      requestType: "Email Request",
      assignedToAccount,
      attachments: attachmentUploads,
      createdAt: email.created_at || new Date().toISOString(),
      createdByAccount: "email-bot",
      completedAt: null,
      completedByAccount: "",
      status: "open",
      inboundEmailId: emailId,
      inboundMessageId: safeText(email.message_id),
      inboundTo: recipients,
      inboundCc: Array.isArray(email.cc) ? email.cc : [],
      inboundFrom: safeText(email.from),
      inboundSubject: safeText(emailView.subject || email.subject),
      inboundReceivedAt: email.created_at || event.created_at || new Date().toISOString(),
    };

    const nextActivityLog = appendActivityEntry(currentSnapshot.activityLog, {
      actor: "email-bot",
      action: "Created request from email",
      scope: "Requests",
      description: `${newRequest.customer} ${newRequest.jobNumber} was added from inbound email${assignedToAccount ? ` for ${assignedToAccount}` : ""}.`,
      details: {
        inboundEmailId: emailId,
        assignedToAccount,
        attachmentCount: attachmentUploads.length,
      },
    });

    await saveSharedSnapshot(
      supabase,
      {
        ...currentSnapshot,
        requests: [newRequest, ...existingRequests],
        activityLog: nextActivityLog,
      },
      "email-bot"
    );

    res.status(200).json({
      ok: true,
      requestId: newRequest.id,
      inboundEmailId: emailId,
      assignedToAccount,
      attachmentCount: attachmentUploads.length,
    });
  } catch (error) {
    console.error("Failed to process inbound Resend email.", error);
    res.status(500).json({
      error: safeText(error?.message) || "Failed to process inbound email.",
    });
  }
}
