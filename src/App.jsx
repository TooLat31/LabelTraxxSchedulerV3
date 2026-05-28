import React, { useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import { isSupabaseConfigured, supabase } from "./lib/supabase";
import { exportWeeklyScheduleWorkbook, importWeeklyScheduleWorkbook } from "./lib/scheduleWorkbook";
import { buildDemoSharedSnapshot, DEMO_DEFAULT_PASSWORD, DEMO_DEFAULT_USERNAME } from "./lib/demoData";

const PRESS_ORDER = ["5.1", "6.1", "1.1", "2.1", "8", "9", "Extra Duties", "Rewind"];
const STORAGE_KEY = "labeltraxx-scheduler-v4";
const SESSION_STORAGE_KEY = "labeltraxx-scheduler-session-v1";
const SHARED_STATE_ROW_ID = "labeltraxx-shared-state";
const LOGIN_SESSION_DURATION_MS = 8 * 60 * 60 * 1000;
const SHARED_SAVE_DEBOUNCE_MS = 700;
const SHARED_REFRESH_INTERVAL_MS = 15000;
const ATTACHMENT_BUCKET = "labeltraxx-attachments";
const ACTIVITY_LOG_LIMIT = 300;
const DEMO_QUERY_PARAM = "demo";
const BASE_TABS = ["Today", "Scheduler", "Notes", "New Request", "Open Requests", "Request History", "Pull Paper Request", "Supplies Request", "Daily Shipment", "Shipment Emails", "Activity Log"];
const ACCESS_MODE_OPTIONS = ["edit", "view"];
const ATTACHMENT_ACCEPT =
  ".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.rtf,.png,.jpg,.jpeg,.zip,.msg,.eml";
const PULL_PAPER_TARGETS = ["Press 5.1", "Press 6.1", "Press 2.1", "Press 1.1", "Digital"];
const ROLE_OPTIONS = ["Management", "Warehouse/Shipper", "Operator"];
const DEFAULT_SHIPMENT_METHODS = ["Skid", "FedEx", "UPS", "LTL", "Customer Pickup"];
const SHIPMENT_QUEUE_WINDOW_OPTIONS = [
  { value: "7", label: "Last 7 days" },
  { value: "30", label: "Last 30 days" },
  { value: "60", label: "Last 60 days" },
  { value: "90", label: "Last 90 days" },
  { value: "all", label: "All finished jobs" },
];
const SCHEDULE_DENSITY_OPTIONS = ["compact", "detailed"];
const SCHEDULE_DENSITY_LABELS = {
  compact: "Compact cards",
  detailed: "Detailed cards",
};

const EMPTY_REQUEST_FORM = {
  jobNumber: "",
  customer: "",
  requestorName: "",
  description: "",
  requestType: "General Request",
  assignedToAccount: "",
};

const EMPTY_PULL_PAPER_FORM = {
  details: "",
  target: PULL_PAPER_TARGETS[0],
};

const EMPTY_SUPPLIES_FORM = {
  details: "",
};

const EMPTY_NOTE_FORM = {
  text: "",
};

const EMPTY_MANUAL_SCHEDULE_FORM = {
  title: "",
  press: PRESS_ORDER[0],
  dayKey: "",
};

const EMPTY_LOGIN_FORM = {
  username: "",
  password: "",
};

const EMPTY_REGISTER_FORM = {
  username: "",
  password: "",
};

const EMPTY_USER_FORM = {
  username: "",
  password: "",
  accessMode: "edit",
  canManageUsers: false,
  tabs: BASE_TABS.filter((tab) => !["New Request", "Open Requests"].includes(tab)),
};

const EMPTY_SHIPMENT_FORM = {
  label: "",
  method: DEFAULT_SHIPMENT_METHODS[0],
  packageCount: "",
  packageType: "Skids",
  totalCost: "",
  billAmount: "",
  notes: "",
};

const EMPTY_EMAIL_FORM = {
  groupId: "",
  recipients: "",
  cc: "",
};

const EMPTY_EMAIL_GROUP_FORM = {
  name: "",
  recipients: "",
  cc: "",
};

const EMPTY_ACTIVITY_FILTERS = {
  actor: "All",
  scope: "All",
  date: "",
};

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function parseNumber(value) {
  const cleaned = safeText(value).replace(/,/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseCurrency(value) {
  const cleaned = safeText(value).replace(/[$,]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  const raw = safeText(value);
  if (!raw || raw === "00/00/00") return null;
  const parts = raw.split(/[\/\-]/).map((part) => part.trim());
  if (parts.length < 3) return null;
  let [monthValue, dayValue, yearValue] = parts;
  const month = Number(monthValue);
  const day = Number(dayValue);
  let year = Number(yearValue);
  if (!Number.isFinite(month) || !Number.isFinite(day) || !Number.isFinite(year)) return null;
  if (year < 100) year += 2000;
  const dt = new Date(year, month - 1, day);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function formatDate(date) {
  if (!date) return "-";
  return new Intl.DateTimeFormat("en-US", {
    month: "numeric",
    day: "numeric",
    year: "2-digit",
  }).format(date);
}

function formatShortDate(date) {
  if (!date) return "";
  return new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    month: "numeric",
    day: "numeric",
  }).format(date);
}

function formatDateTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return new Intl.DateTimeFormat("en-US", {
    month: "numeric",
    day: "numeric",
    year: "2-digit",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(value || 0);
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 KB";
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function startOfWeek(input) {
  const date = new Date(input);
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  date.setHours(0, 0, 0, 0);
  return date;
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function isoDate(date) {
  if (!date) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function todayKey() {
  return isoDate(new Date());
}

function localMiddayIso(date) {
  if (!date) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0, 0).toISOString();
}

function sameDay(dateValue, dayKey) {
  if (!dateValue || !dayKey) return false;
  return isoDate(new Date(dateValue)) === dayKey;
}

function isDemoWorkspaceRequested() {
  if (typeof window === "undefined") return false;
  return new URLSearchParams(window.location.search).get(DEMO_QUERY_PARAM) === "1";
}

function updateDemoWorkspaceUrl(enabled) {
  if (typeof window === "undefined") return;
  const url = new URL(window.location.href);
  if (enabled) {
    url.searchParams.set(DEMO_QUERY_PARAM, "1");
  } else {
    url.searchParams.delete(DEMO_QUERY_PARAM);
  }
  window.history.replaceState({}, "", url.toString());
}

function isWithinLookbackWindow(dateValue, days, referenceDate = new Date()) {
  if (!dateValue || days === "all") return days === "all";
  const parsedDays = Number(days);
  if (!Number.isFinite(parsedDays) || parsedDays <= 0) return true;
  const valueDate = new Date(dateValue);
  if (Number.isNaN(valueDate.getTime())) return false;
  const cutoff = new Date(referenceDate);
  cutoff.setHours(0, 0, 0, 0);
  cutoff.setDate(cutoff.getDate() - parsedDays);
  return valueDate >= cutoff;
}

function comparableUsername(value) {
  return safeText(value).toLowerCase();
}

function getDefaultTabsForRole(role, isAdmin = false) {
  const normalizedRole = normalizeRole(role, isAdmin);
  if (normalizedRole === "Management") return [...BASE_TABS, "User Admin"];
  if (normalizedRole === "Warehouse/Shipper") {
    return BASE_TABS.filter((tab) => !["New Request", "Open Requests"].includes(tab));
  }
  if (normalizedRole === "Operator") return ["Scheduler"];
  return ["Scheduler"];
}

function normalizeAccessMode(mode, role, isAdmin = false) {
  const normalized = safeText(mode).toLowerCase();
  if (ACCESS_MODE_OPTIONS.includes(normalized)) return normalized;
  return normalizeRole(role, isAdmin) === "Operator" ? "view" : "edit";
}

function normalizeUserTabs(tabs, role, isAdmin = false, canManageUsers = false) {
  const next = Array.isArray(tabs)
    ? tabs
        .map((tab) => safeText(tab))
        .filter((tab) => BASE_TABS.includes(tab) || (canManageUsers && tab === "User Admin"))
    : [];
  const fallback = next.length ? Array.from(new Set(next)) : getDefaultTabsForRole(role, isAdmin);
  const filtered = canManageUsers ? fallback : fallback.filter((tab) => tab !== "User Admin");
  if (canManageUsers && !filtered.includes("User Admin")) filtered.push("User Admin");
  return filtered;
}

function normalizeRole(role, isAdmin) {
  const normalized = safeText(role);
  if (ROLE_OPTIONS.includes(normalized)) return normalized;
  return isAdmin ? "Management" : "Warehouse/Shipper";
}

function getUserRole(user) {
  if (user?.canManageUsers) return "Management";
  return normalizeRole(user?.role, user?.isAdmin);
}

function hasManagementAccess(user) {
  return !!user?.canManageUsers || getUserRole(user) === "Management";
}

function getUserAccessMode(user) {
  return normalizeAccessMode(user?.accessMode, user?.role, user?.canManageUsers || user?.isAdmin);
}

function canEdit(user) {
  return getUserAccessMode(user) === "edit";
}

function getVisibleTabs(user) {
  if (!user) return ["Scheduler"];
  return normalizeUserTabs(user.tabs, user.role, user.isAdmin, user.canManageUsers);
}

function canMoveJobs(user) {
  return canEdit(user) && canAccessTab(user, "Scheduler");
}

function canAccessTab(user, tab) {
  return getVisibleTabs(user).includes(tab);
}

function buildDefaultAdmin() {
  return {
    id: "user-admin",
    username: "Admin",
    password: "1234",
    role: "Management",
    accessMode: "edit",
    tabs: [...BASE_TABS, "User Admin"],
    canManageUsers: true,
    isAdmin: true,
    createdAt: new Date().toISOString(),
    createdBy: "system",
  };
}

function normalizeUsers(users) {
  const normalized = Array.isArray(users)
    ? users
        .map((user, index) => ({
          id: user.id || `user-${index + 1}`,
          username: safeText(user.username),
          password: safeText(user.password),
          role: normalizeRole(user.role, user.isAdmin || user.canManageUsers),
          accessMode: normalizeAccessMode(user.accessMode, user.role, user.isAdmin || user.canManageUsers),
          canManageUsers: !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management",
          tabs: normalizeUserTabs(
            user.tabs,
            user.role,
            user.isAdmin || user.canManageUsers,
            !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management"
          ),
          isAdmin: !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management",
          createdAt: user.createdAt || new Date().toISOString(),
          createdBy: user.createdBy || "system",
        }))
        .filter((user) => user.username)
    : [];

  const hasAdmin = normalized.some((user) => comparableUsername(user.username) === "admin");
  if (!hasAdmin) normalized.unshift(buildDefaultAdmin());
  return normalized;
}

function normalizeJobs(jobs) {
  return Array.isArray(jobs)
    ? jobs.map((job) => ({
        ...job,
        press: normalizePressValue(job.press),
        shipByDate: job.shipByDate ? new Date(job.shipByDate) : null,
        entryDate: job.entryDate ? new Date(job.entryDate) : null,
        dueOnSiteDate: job.dueOnSiteDate ? new Date(job.dueOnSiteDate) : null,
        dateDone: job.dateDone ? new Date(job.dateDone) : null,
        holdActive: !!job.holdActive,
        holdNote: safeText(job.holdNote),
      }))
    : [];
}

function buildJobSearchText(job) {
  return `${safeText(job.number)} ${safeText(job.customerName)} ${safeText(job.generalDescr)} ${safeText(job.notes)} ${safeText(job.holdNote)}`.toLowerCase();
}

function sanitizeAttachmentName(name) {
  const cleaned = safeText(name).replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/-+/g, "-");
  return cleaned.replace(/^-|-$/g, "") || "file";
}

function buildAttachmentPath(id, fileName) {
  const stamp = new Date().toISOString().slice(0, 10);
  return `${stamp}/${id}-${sanitizeAttachmentName(fileName)}`;
}

function normalizeAttachment(attachment, index) {
  return {
    id: attachment.id || `attachment-${index + 1}`,
    name: safeText(attachment.name),
    size: parseNumber(attachment.size),
    type: safeText(attachment.type) || "application/octet-stream",
    uploadedAt: attachment.uploadedAt || new Date().toISOString(),
    uploadedBy: safeText(attachment.uploadedBy),
    storagePath: safeText(attachment.storagePath),
    publicUrl: safeText(attachment.publicUrl),
    dataUrl: safeText(attachment.dataUrl),
  };
}

function normalizeAttachments(attachments) {
  return Array.isArray(attachments)
    ? attachments.map((attachment, index) => normalizeAttachment(attachment, index)).filter((attachment) => attachment.name)
    : [];
}

function getAttachmentHref(attachment) {
  return safeText(attachment.publicUrl) || safeText(attachment.dataUrl);
}

function sanitizeEmailHeader(value) {
  return safeText(value).replace(/[\r\n]+/g, " ").trim();
}

function bytesToBase64(bytes) {
  if (!bytes?.length) return "";
  const chunkSize = 0x8000;
  let binary = "";
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...slice);
  }
  return btoa(binary);
}

function utf8ToBase64(text) {
  return bytesToBase64(new TextEncoder().encode(safeText(text)));
}

function wrapBase64(base64Text) {
  return safeText(base64Text).match(/.{1,76}/g)?.join("\r\n") || "";
}

function dataUrlToEmailAttachment(attachment) {
  const dataUrl = safeText(attachment.dataUrl);
  const match = dataUrl.match(/^data:([^;,]+)?(?:;charset=[^;,]+)?;base64,(.*)$/i);
  if (!match) return null;
  return {
    name: sanitizeAttachmentName(attachment.name),
    type: safeText(match[1]) || safeText(attachment.type) || "application/octet-stream",
    base64: safeText(match[2]).replace(/\s+/g, ""),
  };
}

async function resolveEmailAttachment(attachment) {
  const normalized = normalizeAttachment(attachment, 0);
  if (!normalized.name) return null;
  const inlineAttachment = dataUrlToEmailAttachment(normalized);
  if (inlineAttachment) return inlineAttachment;
  const href = getAttachmentHref(normalized);
  if (!href) return null;
  try {
    const response = await fetch(href);
    if (!response.ok) {
      throw new Error(`Attachment request failed with ${response.status}`);
    }
    const bytes = new Uint8Array(await response.arrayBuffer());
    return {
      name: sanitizeAttachmentName(normalized.name),
      type: safeText(normalized.type) || response.headers.get("content-type") || "application/octet-stream",
      base64: bytesToBase64(bytes),
    };
  } catch (error) {
    console.error("Failed to prepare shipment email attachment.", error);
    return null;
  }
}

function buildEmailDraftFile({ recipients, cc, subject, body, attachments }) {
  const boundary = `----=_Part_${makeId("email-boundary")}`;
  const lines = [
    `To: ${sanitizeEmailHeader(recipients)}`,
    cc ? `Cc: ${sanitizeEmailHeader(cc)}` : "",
    `Subject: ${sanitizeEmailHeader(subject)}`,
    "MIME-Version: 1.0",
    "X-Unsent: 1",
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    "",
    "This is a multi-part message in MIME format.",
    `--${boundary}`,
    'Content-Type: text/plain; charset="utf-8"',
    "Content-Transfer-Encoding: base64",
    "",
    wrapBase64(utf8ToBase64(body)),
  ];

  attachments.forEach((attachment) => {
    lines.push(
      `--${boundary}`,
      `Content-Type: ${attachment.type}; name="${attachment.name}"`,
      "Content-Transfer-Encoding: base64",
      `Content-Disposition: attachment; filename="${attachment.name}"`,
      "",
      wrapBase64(attachment.base64)
    );
  });

  lines.push(`--${boundary}--`, "");
  return lines.join("\r\n");
}

function storageUploadsEnabled() {
  return !!(isSupabaseConfigured && supabase);
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

function normalizeActivityLog(log) {
  return Array.isArray(log)
    ? log
        .map((entry, index) => ({
          id: entry.id || `activity-${index + 1}`,
          actor: safeText(entry.actor) || "system",
          action: safeText(entry.action),
          scope: safeText(entry.scope) || "General",
          description: safeText(entry.description),
          createdAt: entry.createdAt || new Date().toISOString(),
          details: entry.details && typeof entry.details === "object" ? entry.details : {},
        }))
        .filter((entry) => entry.action || entry.description)
    : [];
}

function normalizeRequests(requests) {
  return Array.isArray(requests)
    ? requests.map((request) => ({
        ...request,
        attachments: normalizeAttachments(request.attachments),
        createdByAccount: request.createdByAccount || "",
        completedByAccount: request.completedByAccount || "",
        requestType: safeText(request.requestType) || "General Request",
        assignedToAccount: safeText(request.assignedToAccount),
      }))
    : [];
}

function normalizePullPaperRequests(requests) {
  return Array.isArray(requests) ? requests : [];
}

function normalizeNotes(notes) {
  return Array.isArray(notes)
    ? notes
        .map((note, index) => ({
          id: note.id || `note-${index + 1}`,
          ownerUsername: safeText(note.ownerUsername || note.username),
          text: safeText(note.text),
          completed: !!note.completed,
          createdAt: note.createdAt || new Date().toISOString(),
          completedAt: note.completedAt || "",
        }))
        .filter((note) => note.ownerUsername && note.text)
    : [];
}

function normalizeRegistrationRequests(requests) {
  return Array.isArray(requests)
    ? requests
        .map((request, index) => ({
          id: request.id || `registration-${index + 1}`,
          username: safeText(request.username),
          password: safeText(request.password),
          status: safeText(request.status) || "pending",
          createdAt: request.createdAt || new Date().toISOString(),
          createdBy: request.createdBy || safeText(request.username),
          approvedAt: request.approvedAt || "",
          approvedBy: request.approvedBy || "",
          deniedAt: request.deniedAt || "",
          deniedBy: request.deniedBy || "",
        }))
        .filter((request) => request.username && request.password)
    : [];
}

function normalizeSuppliesRequests(requests) {
  return Array.isArray(requests)
    ? requests.map((request) => ({
        ...request,
        attachments: normalizeAttachments(request.attachments),
      }))
    : [];
}

function normalizeAssignments(assignments) {
  return Array.isArray(assignments)
    ? assignments
        .filter((assignment) => assignment.kind !== "rewind")
        .map((assignment, index) => ({
          id: assignment.id || `assignment-${index + 1}`,
          ...assignment,
          manualTitle: safeText(assignment.manualTitle),
          kind: safeText(assignment.kind) || "press",
          status: safeText(assignment.status) || "scheduled",
          laneOrder: Number.isFinite(Number(assignment.laneOrder)) ? Number(assignment.laneOrder) : null,
        }))
    : [];
}

function normalizeShipmentGroups(groups) {
  return Array.isArray(groups)
    ? groups.map((group) => ({
        ...group,
        attachments: normalizeAttachments(group.attachments),
        billAmount: parseCurrency(group.billAmount),
      }))
    : [];
}

function normalizeShipmentMethods(methods) {
  const next = Array.isArray(methods)
    ? methods.map((method) => safeText(method)).filter(Boolean)
    : [];
  return next.length ? Array.from(new Set(next)) : [...DEFAULT_SHIPMENT_METHODS];
}

function normalizeShipmentEmailLogs(logs) {
  return Array.isArray(logs)
    ? logs
        .map((log, index) => ({
          id: log.id || `email-${index + 1}`,
          shipDate: safeText(log.shipDate),
          recipients: safeText(log.recipients),
          cc: safeText(log.cc),
          subject: safeText(log.subject),
          body: safeText(log.body),
          jobCount: parseNumber(log.jobCount),
          totalCost: parseCurrency(log.totalCost),
          totalBill: parseCurrency(log.totalBill),
          methods: Array.isArray(log.methods) ? log.methods.map((method) => safeText(method)).filter(Boolean) : [],
          createdAt: log.createdAt || new Date().toISOString(),
          createdBy: safeText(log.createdBy),
        }))
        .filter((log) => log.shipDate)
    : [];
}

function normalizeShipmentEmailGroups(groups) {
  return Array.isArray(groups)
    ? groups
        .map((group, index) => ({
          id: group.id || `email-group-${index + 1}`,
          name: safeText(group.name),
          recipients: safeText(group.recipients),
          cc: safeText(group.cc),
        }))
        .filter((group) => group.name && (group.recipients || group.cc))
    : [];
}

function normalizePressOperators(operators) {
  const source = operators && typeof operators === "object" ? operators : {};
  return Object.fromEntries(
    Object.entries(source)
      .map(([key, value]) => [safeText(key), safeText(value)])
      .filter(([key, value]) => key && value)
  );
}

function normalizePressDuties(duties) {
  const source = duties && typeof duties === "object" ? duties : {};
  return Object.fromEntries(
    Object.entries(source)
      .map(([key, value]) => [safeText(key), safeText(value)])
      .filter(([key, value]) => key && value)
  );
}

function defaultSharedSnapshot() {
  return {
    jobs: [],
    assignments: [],
    pressOperators: {},
    pressDuties: {},
    requests: [],
    pullPaperRequests: [],
    notes: [],
    registrationRequests: [],
    suppliesRequests: [],
    shipmentGroups: [],
    shipmentEmailLogs: [],
    shipmentEmailGroups: [],
    activityLog: [],
    shipmentMethods: [...DEFAULT_SHIPMENT_METHODS],
    users: [buildDefaultAdmin()],
    weekStart: startOfWeek(new Date()).toISOString(),
  };
}

function normalizeSharedSnapshot(snapshot) {
  const source = snapshot && typeof snapshot === "object" ? snapshot : {};
  return {
    jobs: normalizeJobs(source.jobs),
    assignments: normalizeAssignments(source.assignments),
    pressOperators: normalizePressOperators(source.pressOperators),
    pressDuties: normalizePressDuties(source.pressDuties),
    requests: normalizeRequests(source.requests),
    pullPaperRequests: normalizePullPaperRequests(source.pullPaperRequests),
    notes: normalizeNotes(source.notes),
    registrationRequests: normalizeRegistrationRequests(source.registrationRequests),
    suppliesRequests: normalizeSuppliesRequests(source.suppliesRequests),
    shipmentGroups: normalizeShipmentGroups(source.shipmentGroups),
    shipmentEmailLogs: normalizeShipmentEmailLogs(source.shipmentEmailLogs),
    shipmentEmailGroups: normalizeShipmentEmailGroups(source.shipmentEmailGroups),
    activityLog: normalizeActivityLog(source.activityLog),
    shipmentMethods: normalizeShipmentMethods(source.shipmentMethods),
    users: normalizeUsers(source.users),
    weekStart: source.weekStart ? new Date(source.weekStart) : startOfWeek(new Date()),
  };
}

function buildSharedSnapshot(state) {
  return {
    jobs: state.jobs,
    assignments: state.assignments,
    pressOperators: state.pressOperators,
    pressDuties: state.pressDuties,
    requests: state.requests,
    pullPaperRequests: state.pullPaperRequests,
    notes: state.notes,
    registrationRequests: state.registrationRequests,
    suppliesRequests: state.suppliesRequests,
    shipmentGroups: state.shipmentGroups,
    shipmentEmailLogs: state.shipmentEmailLogs,
    shipmentEmailGroups: state.shipmentEmailGroups,
    activityLog: state.activityLog,
    shipmentMethods: state.shipmentMethods,
    users: state.users,
    weekStart: state.weekStart.toISOString(),
  };
}

function readFileAsDataUrlAttachment(file, uploadedBy, id = makeId("att")) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () =>
      resolve({
        id,
        name: file.name,
        size: file.size,
        type: file.type || "application/octet-stream",
        uploadedAt: new Date().toISOString(),
        uploadedBy,
        storagePath: "",
        publicUrl: "",
        dataUrl: String(reader.result || ""),
      });
    reader.onerror = () => reject(reader.error || new Error("Failed to read file."));
    reader.readAsDataURL(file);
  });
}

async function buildAttachment(file, uploadedBy) {
  const id = makeId("att");
  const baseAttachment = {
    id,
    name: file.name,
    size: file.size,
    type: file.type || "application/octet-stream",
    uploadedAt: new Date().toISOString(),
    uploadedBy,
    storagePath: "",
    publicUrl: "",
    dataUrl: "",
  };

  if (storageUploadsEnabled()) {
    try {
      const storagePath = buildAttachmentPath(id, file.name);
      const { error } = await supabase.storage.from(ATTACHMENT_BUCKET).upload(storagePath, file, {
        upsert: false,
        contentType: baseAttachment.type,
      });
      if (!error) {
        const { data } = supabase.storage.from(ATTACHMENT_BUCKET).getPublicUrl(storagePath);
        return {
          ...baseAttachment,
          storagePath,
          publicUrl: safeText(data?.publicUrl),
        };
      }
      console.warn("Falling back to inline attachment storage.", error);
    } catch (error) {
      console.error("Failed to upload attachment to Supabase storage.", error);
    }
  }

  return readFileAsDataUrlAttachment(file, uploadedBy, id);
}

async function buildAttachments(fileList, uploadedBy) {
  const files = Array.from(fileList || []);
  if (!files.length) return [];
  return Promise.all(files.map((file) => buildAttachment(file, uploadedBy)));
}

async function deleteStoredAttachments(attachments) {
  const removablePaths = normalizeAttachments(attachments)
    .map((attachment) => safeText(attachment.storagePath))
    .filter(Boolean);
  if (!removablePaths.length || !storageUploadsEnabled()) return;
  try {
    const { error } = await supabase.storage.from(ATTACHMENT_BUCKET).remove(removablePaths);
    if (error) {
      console.error("Failed to delete attachment from Supabase storage.", error);
    }
  } catch (error) {
    console.error("Failed to delete attachment from Supabase storage.", error);
  }
}

function looksLikeJobRow(parts) {
  if (parts.length < 2) return false;
  const press = safeText(parts[0]);
  const number = safeText(parts[1]);
  return /^\d+(?:\.\d+)?$/.test(press) && /^\d+$/.test(number);
}

function normalizeJobStatus(status) {
  const value = safeText(status).toLowerCase();
  if (!value) return "open";
  if (value === "done" || value === "closed" || value === "finished" || value === "complete") return "closed";
  return "open";
}

function normalizePressValue(value) {
  const normalized = safeText(value);
  if (!normalized) return "";
  if (normalized === "1") return "1.1";
  if (normalized === "2") return "2.1";
  return normalized;
}

function formatPressLabel(press) {
  const normalized = safeText(press);
  if (!normalized) return "-";
  return normalized === "Extra Duties" ? normalized : `Press ${normalized}`;
}

function isReleaseJob(job) {
  return comparableUsername(job?.priority).includes("release");
}

function parseLabelTraxxText(text) {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const headerLine = lines.find(
    (line) => line.includes("Press\t") && line.includes("Number\t") && line.includes("CustomerName")
  );
  if (!headerLine) return [];

  const startIndex = lines.indexOf(headerLine);
  const headers = headerLine.split("\t");
  const headerCount = headers.length;
  const notesIndex = headers.findIndex((header) => safeText(header).toLowerCase() === "notes");
  const dateDoneIndex = headers.findIndex((header) => safeText(header).toLowerCase() === "datedone");
  const rows = [];
  let current = null;

  for (let index = startIndex + 1; index < lines.length; index += 1) {
    const line = lines[index];
    if (!line.trim()) {
      if (current) {
        const targetIndex = notesIndex >= 0 ? notesIndex : headerCount - 1;
        current[targetIndex] = `${current[targetIndex] || ""}\n`;
      }
      continue;
    }

    const parts = line.split("\t");
    if (looksLikeJobRow(parts)) {
      if (current) rows.push(current);
      current = Array.from({ length: headerCount }, (_, itemIndex) => parts[itemIndex] ?? "");
      if (parts.length > headerCount) {
        current[headerCount - 1] = parts.slice(headerCount - 1).join("\t");
      }
      continue;
    }

    if (current) {
      const maybeDateDone = dateDoneIndex >= 0 ? safeText(parts[parts.length - 1]) : "";
      if (dateDoneIndex >= 0 && notesIndex >= 0 && parts.length >= 2 && parseDate(maybeDateDone)) {
        const noteText = safeText(parts.slice(0, -1).join("\t"));
        current[notesIndex] = `${current[notesIndex] || ""}${current[notesIndex] ? "\n" : ""}${noteText}`;
        current[dateDoneIndex] = maybeDateDone;
        continue;
      }
      const targetIndex = notesIndex >= 0 ? notesIndex : headerCount - 1;
      current[targetIndex] = `${current[targetIndex] || ""}${current[targetIndex] ? "\n" : ""}${line}`;
    }
  }

  if (current) rows.push(current);

  return rows
    .map((row) => {
      const record = Object.fromEntries(headers.map((header, index) => [header, safeText(row[index])]));
      const stockParts = [safeText(record.StockNum2), safeText(record.StockNum1)].filter(Boolean);
      const importedStatus = safeText(record.Status) || safeText(record.TicketStatus) || "Open";
      return {
        id: safeText(record.Number),
        press: normalizePressValue(record.Press),
        number: safeText(record.Number),
        customerName: safeText(record.CustomerName),
        generalDescr: safeText(record.GeneralDescr),
        custPoNum: safeText(record.CustPONum),
        priority: safeText(record.Priority),
        shipByDate: parseDate(record.Ship_by_Date),
        entryDate: parseDate(record.EntryDate),
        dueOnSiteDate: parseDate(record.Due_on_Site_Date),
        stockNum2: safeText(record.StockNum2),
        stockNum1: safeText(record.StockNum1),
        stockDisplay: stockParts.join(" / "),
        ticketStatus: importedStatus,
        normalizedStatus: normalizeJobStatus(importedStatus),
        mainTool: safeText(record.MainTool),
        toolNo2: safeText(record.ToolNo2),
        ticQuantity: parseNumber(record.TicQuantity),
        estFootage: parseNumber(record.EstFootage),
        estPressTime: parseNumber(record.EstPressTime),
        notes: safeText(record.Notes),
        dateDone: parseDate(record.DateDone),
        raw: record,
      };
    })
    .filter((job) => job.number);
}

function priorityTone(priority) {
  const value = safeText(priority).toLowerCase();
  if (value.includes("high")) return "bg-amber-100 text-amber-900 border-amber-300";
  if (value.includes("release")) return "bg-stone-200 text-stone-800 border-stone-300";
  if (value.includes("inventory")) return "bg-lime-100 text-lime-900 border-lime-300";
  if (value.includes("digital")) return "bg-slate-200 text-slate-800 border-slate-300";
  if (value.includes("test")) return "bg-neutral-200 text-neutral-800 border-neutral-300";
  return "bg-emerald-100 text-emerald-900 border-emerald-300";
}

function statusTone(status) {
  const value = safeText(status).toLowerCase();
  if (value === "finished") return "bg-stone-900 text-stone-50";
  if (value === "ship") return "bg-emerald-800 text-white";
  if (value === "scheduled") return "bg-stone-700 text-stone-50";
  if (value === "done") return "bg-emerald-100 text-emerald-900";
  if (value === "note") return "bg-amber-100 text-amber-900";
  if (value === "open") return "bg-stone-200 text-stone-800";
  return "bg-stone-100 text-stone-700";
}

function syncTone(status) {
  const value = safeText(status).toLowerCase();
  if (value.includes("live")) return "bg-emerald-100 text-emerald-900";
  if (value.includes("demo")) return "bg-sky-100 text-sky-900";
  if (value.includes("saving") || value.includes("connecting")) return "bg-amber-100 text-amber-900";
  if (value.includes("error")) return "bg-rose-100 text-rose-900";
  return "bg-stone-100 text-stone-700";
}

function buildWeekColumns(weekStart) {
  return Array.from({ length: 5 }, (_, index) => {
    const date = addDays(weekStart, index);
    return {
      key: isoDate(date),
      date,
      label: new Intl.DateTimeFormat("en-US", { weekday: "long" }).format(date),
    };
  });
}

function csvEscape(value) {
  const text = safeText(value).replace(/"/g, '""');
  return `"${text}"`;
}

function makeId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function makeDragPayload(payload) {
  return JSON.stringify(payload);
}

function parseDragPayload(raw) {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function dateSortValue(value) {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return Number.isFinite(time) ? time : 0;
}

function downloadFile(name, content, type = "text/plain;charset=utf-8") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function snapshotJobForShipment(job, finishMeta) {
  return {
    id: job.id,
    number: job.number,
    customerName: job.customerName,
    generalDescr: job.generalDescr,
    press: job.press,
    ticQuantity: job.ticQuantity,
    shipByDate: job.shipByDate ? job.shipByDate.toISOString() : null,
    finishedAt: finishMeta?.finishedAt || null,
    finishedBy: finishMeta?.finishedBy || "",
  };
}

function effectiveFinishedShipDate(finishMeta) {
  if (!finishMeta) return "";
  if (finishMeta.shipDate) return finishMeta.shipDate;
  if (finishMeta.dayKey) return finishMeta.dayKey;
  return finishMeta.finishedAt ? isoDate(new Date(finishMeta.finishedAt)) : "";
}

function getShipmentItems(group, jobMap, finishedMetaByJobId) {
  if (Array.isArray(group.jobItems) && group.jobItems.length) return group.jobItems;
  return (group.jobIds || [])
    .map((jobId) => {
      const job = jobMap.get(jobId);
      if (!job) return null;
      return snapshotJobForShipment(job, finishedMetaByJobId.get(jobId));
    })
    .filter(Boolean);
}

function deriveVisibleJobState(jobId, activePressJobIds, finishedJobIds) {
  if (finishedJobIds.has(jobId)) return "finished";
  if (activePressJobIds.has(jobId)) return "scheduled";
  return "open";
}

function pressOperatorKey(dayKey, press) {
  return `${safeText(dayKey)}::${safeText(press)}`;
}

function canUserViewRequest(user, request) {
  if (!request) return false;
  if (hasManagementAccess(user)) return true;
  const username = safeText(user?.username);
  if (!username) return false;
  const assignedTo = safeText(request.assignedToAccount);
  if (!assignedTo) return true;
  return comparableUsername(assignedTo) === comparableUsername(username) || comparableUsername(request.createdByAccount) === comparableUsername(username);
}

function SchedulerApp() {
  const [isReady, setIsReady] = useState(false);
  const [workspaceMode, setWorkspaceMode] = useState(() => (isDemoWorkspaceRequested() ? "demo" : "live"));
  const [jobs, setJobs] = useState([]);
  const [assignments, setAssignments] = useState([]);
  const [pressOperators, setPressOperators] = useState({});
  const [pressDuties, setPressDuties] = useState({});
  const [requests, setRequests] = useState([]);
  const [pullPaperRequests, setPullPaperRequests] = useState([]);
  const [notes, setNotes] = useState([]);
  const [registrationRequests, setRegistrationRequests] = useState([]);
  const [suppliesRequests, setSuppliesRequests] = useState([]);
  const [shipmentGroups, setShipmentGroups] = useState([]);
  const [shipmentEmailLogs, setShipmentEmailLogs] = useState([]);
  const [shipmentEmailGroups, setShipmentEmailGroups] = useState([]);
  const [activityLog, setActivityLog] = useState([]);
  const [shipmentMethods, setShipmentMethods] = useState([...DEFAULT_SHIPMENT_METHODS]);
  const [users, setUsers] = useState([buildDefaultAdmin()]);
  const [currentUsername, setCurrentUsername] = useState("");
  const [weekStart, setWeekStart] = useState(startOfWeek(new Date()));
  const [search, setSearch] = useState("");
  const [unscheduledSearch, setUnscheduledSearch] = useState("");
  const [queueStatusFilter, setQueueStatusFilter] = useState("All");
  const [queuePressFilter, setQueuePressFilter] = useState("All");
  const [queueScheduleFilter, setQueueScheduleFilter] = useState("All");
  const [selectedJobId, setSelectedJobId] = useState(null);
  const [activeTab, setActiveTab] = useState("Scheduler");
  const [requestForm, setRequestForm] = useState(EMPTY_REQUEST_FORM);
  const [pullPaperForm, setPullPaperForm] = useState(EMPTY_PULL_PAPER_FORM);
  const [suppliesForm, setSuppliesForm] = useState(EMPTY_SUPPLIES_FORM);
  const [noteForm, setNoteForm] = useState(EMPTY_NOTE_FORM);
  const [requestDraftAttachments, setRequestDraftAttachments] = useState([]);
  const [suppliesDraftAttachments, setSuppliesDraftAttachments] = useState([]);
  const [selectedShipDate, setSelectedShipDate] = useState(todayKey());
  const [shipDateDraft, setShipDateDraft] = useState(todayKey());
  const [selectedShipmentJobs, setSelectedShipmentJobs] = useState([]);
  const [selectedShipQueueJobs, setSelectedShipQueueJobs] = useState([]);
  const [shipQueueWindow, setShipQueueWindow] = useState("30");
  const [shipmentForm, setShipmentForm] = useState({ ...EMPTY_SHIPMENT_FORM, shipDate: todayKey() });
  const [shipmentDraftAttachments, setShipmentDraftAttachments] = useState([]);
  const [newShipmentMethod, setNewShipmentMethod] = useState("");
  const [shipmentEmailForm, setShipmentEmailForm] = useState(EMPTY_EMAIL_FORM);
  const [shipmentEmailGroupForm, setShipmentEmailGroupForm] = useState(EMPTY_EMAIL_GROUP_FORM);
  const [loginForm, setLoginForm] = useState(EMPTY_LOGIN_FORM);
  const [loginError, setLoginError] = useState("");
  const [authView, setAuthView] = useState("login");
  const [registerForm, setRegisterForm] = useState(EMPTY_REGISTER_FORM);
  const [registerError, setRegisterError] = useState("");
  const [registerSuccess, setRegisterSuccess] = useState("");
  const [sessionExpiresAt, setSessionExpiresAt] = useState("");
  const [userForm, setUserForm] = useState(EMPTY_USER_FORM);
  const [userPasswordDrafts, setUserPasswordDrafts] = useState({});
  const [userUsernameDrafts, setUserUsernameDrafts] = useState({});
  const [requestHistoryFilterDate, setRequestHistoryFilterDate] = useState("");
  const [shipmentHistoryFilterDate, setShipmentHistoryFilterDate] = useState("");
  const [shipmentEmailHistoryFilterDate, setShipmentEmailHistoryFilterDate] = useState("");
  const [activityFilters, setActivityFilters] = useState(EMPTY_ACTIVITY_FILTERS);
  const [queuePriorityFilters, setQueuePriorityFilters] = useState([]);
  const [visiblePressFilters, setVisiblePressFilters] = useState([...PRESS_ORDER]);
  const [scheduleCardDensity, setScheduleCardDensity] = useState("compact");
  const [hideEmptyPresses, setHideEmptyPresses] = useState(false);
  const [boardFocusMode, setBoardFocusMode] = useState(false);
  const [manualScheduleForm, setManualScheduleForm] = useState(EMPTY_MANUAL_SCHEDULE_FORM);
  const [locationSearch, setLocationSearch] = useState("");
  const [pickedUpItem, setPickedUpItem] = useState(null);
  const [syncStatus, setSyncStatus] = useState(isSupabaseConfigured ? "Connecting..." : "Local only");
  const [lastSyncAt, setLastSyncAt] = useState("");
  const jobDetailsRef = useRef(null);
  const lastSharedSnapshotRef = useRef("");
  const lastLocalSharedChangeRef = useRef(0);
  const saveTimerRef = useRef(null);
  const deferredSearch = useDeferredValue(search);
  const deferredUnscheduledSearch = useDeferredValue(unscheduledSearch);
  const deferredLocationSearch = useDeferredValue(locationSearch);

  function applySharedStateSnapshot(snapshot) {
    const normalized = normalizeSharedSnapshot(snapshot);
    setJobs(normalized.jobs);
    setAssignments(normalized.assignments);
    setPressOperators(normalized.pressOperators);
    setPressDuties(normalized.pressDuties);
    setRequests(normalized.requests);
    setPullPaperRequests(normalized.pullPaperRequests);
    setNotes(normalized.notes);
    setRegistrationRequests(normalized.registrationRequests);
    setSuppliesRequests(normalized.suppliesRequests);
    setShipmentGroups(normalized.shipmentGroups);
    setShipmentEmailLogs(normalized.shipmentEmailLogs);
    setShipmentEmailGroups(normalized.shipmentEmailGroups);
    setActivityLog(normalized.activityLog);
    setShipmentMethods(normalized.shipmentMethods);
    setUsers(normalized.users);
    setWeekStart(normalized.weekStart);
    return normalized;
  }

  async function fetchLatestSharedState({ force = false } = {}) {
    if (!isSupabaseConfigured || !supabase) return null;
    if (!force && Date.now() - lastLocalSharedChangeRef.current < SHARED_SAVE_DEBOUNCE_MS + 400) return null;

    const { data, error } = await supabase
      .from("app_state")
      .select("payload, updated_at")
      .eq("id", SHARED_STATE_ROW_ID)
      .maybeSingle();

    if (error) throw error;
    if (!data?.payload) return null;

    const normalized = normalizeSharedSnapshot(data.payload);
    const digest = JSON.stringify(buildSharedSnapshot(normalized));
    if (digest !== lastSharedSnapshotRef.current) {
      applySharedStateSnapshot(data.payload);
      lastSharedSnapshotRef.current = digest;
    }

    setSyncStatus("Live sync");
    setLastSyncAt(data.updated_at || new Date().toISOString());
    return normalized;
  }

  useEffect(() => {
    let isCancelled = false;

    async function hydrate() {
      let saved = {};
      try {
        saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
        const session = JSON.parse(localStorage.getItem(SESSION_STORAGE_KEY) || "{}");
        const demoRequested = isDemoWorkspaceRequested();
        const sessionWorkspaceMode = safeText(session.workspaceMode) || "live";
        const sessionUsername = safeText(session.currentUsername);
        const storedExpiresAt = safeText(session.expiresAt);
        const expiresAtValue = storedExpiresAt ? new Date(storedExpiresAt).getTime() : 0;
        const isSessionActive = sessionUsername && Number.isFinite(expiresAtValue) && expiresAtValue > Date.now();

        if (demoRequested) {
          const demoSnapshot = buildDemoSharedSnapshot();
          const normalized = applySharedStateSnapshot(demoSnapshot);
          const digest = JSON.stringify(buildSharedSnapshot(normalized));
          lastSharedSnapshotRef.current = digest;
          if (!isCancelled) {
            setWorkspaceMode("demo");
            setSyncStatus("Demo mode");
            setLastSyncAt("");
          }
          const nextDemoUsername =
            isSessionActive && sessionWorkspaceMode === "demo"
              ? normalized.users.find((user) => comparableUsername(user.username) === comparableUsername(sessionUsername))?.username ||
                DEMO_DEFAULT_USERNAME
              : DEMO_DEFAULT_USERNAME;
          if (!isCancelled) {
            setCurrentUsername(nextDemoUsername);
            setSessionExpiresAt(new Date(Date.now() + LOGIN_SESSION_DURATION_MS).toISOString());
          }
          return;
        }

        let sharedSnapshot = saved;
        if (!isCancelled) {
          setWorkspaceMode("live");
        }

        if (isSupabaseConfigured && supabase) {
          setSyncStatus("Connecting...");
          const { data, error } = await supabase
            .from("app_state")
            .select("payload, updated_at")
            .eq("id", SHARED_STATE_ROW_ID)
            .maybeSingle();

          if (error) throw error;

          if (data?.payload) {
            sharedSnapshot = data.payload;
            if (!isCancelled) {
              setSyncStatus("Live sync");
              setLastSyncAt(data.updated_at || new Date().toISOString());
            }
          } else {
            const seedSnapshot =
              Object.keys(saved || {}).length > 0 ? buildSharedSnapshot(normalizeSharedSnapshot(saved)) : defaultSharedSnapshot();
            const { error: seedError } = await supabase.from("app_state").upsert({
              id: SHARED_STATE_ROW_ID,
              payload: seedSnapshot,
              updated_by: safeText(session.currentUsername) || "system",
            });
            if (seedError) throw seedError;
            sharedSnapshot = seedSnapshot;
            if (!isCancelled) {
              setSyncStatus("Live sync");
              setLastSyncAt(new Date().toISOString());
            }
          }
        }

        if (isCancelled) return;

        const normalized = applySharedStateSnapshot(sharedSnapshot);
        const digest = JSON.stringify(buildSharedSnapshot(normalized));
        lastSharedSnapshotRef.current = digest;
        if (isSessionActive) {
          const match = normalized.users.find(
            (user) => comparableUsername(user.username) === comparableUsername(sessionUsername)
          );
          if (match) {
            setCurrentUsername(match.username);
            setSessionExpiresAt(storedExpiresAt);
          }
        } else if (sessionUsername || storedExpiresAt) {
          localStorage.setItem(
            SESSION_STORAGE_KEY,
            JSON.stringify({
              currentUsername: "",
              expiresAt: "",
            })
          );
        }
      } catch (error) {
        console.error("Failed to load shared scheduler state.", error);
        const fallbackSource = Object.keys(saved || {}).length ? saved : defaultSharedSnapshot();
        const fallback = normalizeSharedSnapshot(fallbackSource);
        applySharedStateSnapshot(fallback);
        lastSharedSnapshotRef.current = JSON.stringify(buildSharedSnapshot(fallback));
        setSyncStatus(isSupabaseConfigured ? "Sync error" : "Local only");
      } finally {
        if (!isCancelled) setIsReady(true);
      }
    }

    hydrate();

    return () => {
      isCancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!isReady) return;
    if (workspaceMode === "demo") {
      setSyncStatus("Demo mode");
      return;
    }
    const sharedSnapshot = buildSharedSnapshot({
      jobs,
      assignments,
      pressOperators,
      pressDuties,
      requests,
      pullPaperRequests,
      notes,
      registrationRequests,
      suppliesRequests,
      shipmentGroups,
      shipmentEmailLogs,
      shipmentEmailGroups,
      activityLog,
      shipmentMethods,
      users,
      weekStart,
    });
    lastLocalSharedChangeRef.current = Date.now();
    window.clearTimeout(saveTimerRef.current);
    if (isSupabaseConfigured && supabase) {
      setSyncStatus("Saving...");
    }

    saveTimerRef.current = window.setTimeout(async () => {
      const digest = JSON.stringify(sharedSnapshot);
      localStorage.setItem(STORAGE_KEY, digest);

      if (digest === lastSharedSnapshotRef.current) {
        if (!isSupabaseConfigured || !supabase) {
          setSyncStatus("Local only");
        } else {
          setSyncStatus("Live sync");
        }
        return;
      }

      lastSharedSnapshotRef.current = digest;

      if (!isSupabaseConfigured || !supabase) {
        setSyncStatus("Local only");
        return;
      }

      const { error } = await supabase.from("app_state").upsert({
        id: SHARED_STATE_ROW_ID,
        payload: sharedSnapshot,
        updated_by: currentUsername || "system",
      });

      if (error) {
        console.error("Failed to save shared scheduler state.", error);
        setSyncStatus("Sync error");
        return;
      }

      setSyncStatus("Live sync");
      setLastSyncAt(new Date().toISOString());
    }, SHARED_SAVE_DEBOUNCE_MS);

    return () => {
      window.clearTimeout(saveTimerRef.current);
    };
  }, [activityLog, assignments, currentUsername, isReady, jobs, notes, pressDuties, pressOperators, pullPaperRequests, registrationRequests, requests, shipmentEmailGroups, shipmentEmailLogs, shipmentGroups, shipmentMethods, suppliesRequests, users, weekStart, workspaceMode]);

  useEffect(() => {
    localStorage.setItem(
      SESSION_STORAGE_KEY,
      JSON.stringify({
        currentUsername,
        expiresAt: currentUsername ? sessionExpiresAt : "",
        workspaceMode,
      })
    );
  }, [currentUsername, sessionExpiresAt, workspaceMode]);

  const currentUser = useMemo(
    () =>
      users.find((user) => comparableUsername(user.username) === comparableUsername(currentUsername)) || null,
    [currentUsername, users]
  );

  const currentUserRole = useMemo(() => getUserRole(currentUser), [currentUser]);
  const currentUserAccessMode = useMemo(() => getUserAccessMode(currentUser), [currentUser]);
  const userCanEdit = useMemo(() => canEdit(currentUser), [currentUser]);
  const userCanManageUsers = useMemo(() => hasManagementAccess(currentUser), [currentUser]);
  const userCanMoveJobs = useMemo(() => canMoveJobs(currentUser), [currentUser]);

  function recordActivity(action, scope, description, details = {}) {
    setActivityLog((current) =>
      appendActivityEntry(current, {
        actor: currentUser?.username || "system",
        action,
        scope,
        description,
        details,
      })
    );
  }

  useEffect(() => {
    if (currentUser) return;
    if (!currentUsername) return;
    setCurrentUsername("");
    setSessionExpiresAt("");
  }, [currentUser, currentUsername]);

  useEffect(() => {
    if (!currentUsername || !sessionExpiresAt) return;
    const expiresAtValue = new Date(sessionExpiresAt).getTime();
    if (!Number.isFinite(expiresAtValue)) {
      setCurrentUsername("");
      setSessionExpiresAt("");
      return;
    }
    const timeoutMs = expiresAtValue - Date.now();
    if (timeoutMs <= 0) {
      setCurrentUsername("");
      setSessionExpiresAt("");
      return;
    }
    const timer = window.setTimeout(() => {
      setCurrentUsername("");
      setSessionExpiresAt("");
    }, timeoutMs);
    return () => window.clearTimeout(timer);
  }, [currentUsername, sessionExpiresAt]);

  useEffect(() => {
    if (!currentUser) return;
    if (canAccessTab(currentUser, activeTab)) return;
    setActiveTab("Scheduler");
  }, [activeTab, currentUser]);

  useEffect(() => {
    if (!isReady || workspaceMode === "demo" || !isSupabaseConfigured || !supabase) return undefined;

    let isCancelled = false;

    const refreshFromServer = async (force = false) => {
      try {
        await fetchLatestSharedState({ force });
      } catch (error) {
        if (isCancelled) return;
        console.error("Failed to refresh shared scheduler state.", error);
        setSyncStatus("Sync error");
      }
    };

    const handleFocus = () => {
      refreshFromServer(true);
    };

    const handleVisibilityChange = () => {
      if (document.visibilityState === "visible") {
        refreshFromServer(true);
      }
    };

    const intervalId = window.setInterval(() => {
      if (document.visibilityState === "visible") {
        refreshFromServer();
      }
    }, SHARED_REFRESH_INTERVAL_MS);

    window.addEventListener("focus", handleFocus);
    document.addEventListener("visibilitychange", handleVisibilityChange);

    return () => {
      isCancelled = true;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", handleFocus);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, [isReady, workspaceMode]);

  useEffect(() => {
    if (!isReady || workspaceMode === "demo" || !isSupabaseConfigured || !supabase) return undefined;

    const channel = supabase
      .channel("labeltraxx-shared-state")
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "app_state",
          filter: `id=eq.${SHARED_STATE_ROW_ID}`,
        },
        (payload) => {
          const nextPayload = payload.new?.payload;
          if (!nextPayload) return;
          const digest = JSON.stringify(nextPayload);
          if (digest === lastSharedSnapshotRef.current) {
            setSyncStatus("Live sync");
            setLastSyncAt(payload.new?.updated_at || new Date().toISOString());
            return;
          }

          const normalized = applySharedStateSnapshot(nextPayload);
          lastSharedSnapshotRef.current = JSON.stringify(buildSharedSnapshot(normalized));
          setSyncStatus("Live sync");
          setLastSyncAt(payload.new?.updated_at || new Date().toISOString());
        }
      )
      .subscribe((status) => {
        if (status === "SUBSCRIBED") {
          setSyncStatus("Live sync");
          return;
        }
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          setSyncStatus("Sync error");
          fetchLatestSharedState({ force: true }).catch((error) => {
            console.error("Failed to recover shared scheduler state after realtime error.", error);
          });
        }
      });

    return () => {
      supabase.removeChannel(channel);
    };
  }, [isReady, workspaceMode]);

  useEffect(() => {
    if (!jobs.some((job) => job.id === selectedJobId)) setSelectedJobId(null);
  }, [jobs, selectedJobId]);

  useEffect(() => {
    setShipmentForm((current) => ({ ...current, shipDate: selectedShipDate }));
    setShipDateDraft(selectedShipDate);
    setSelectedShipmentJobs([]);
  }, [selectedShipDate]);

  useEffect(() => {
    if (!shipmentMethods.length) return;
    setShipmentForm((current) =>
      shipmentMethods.includes(current.method) ? current : { ...current, method: shipmentMethods[0] }
    );
  }, [shipmentMethods]);

  const tabs = useMemo(
    () => [...BASE_TABS, "User Admin"].filter((tab) => canAccessTab(currentUser, tab)),
    [currentUser]
  );

  const weekColumns = useMemo(() => buildWeekColumns(weekStart), [weekStart]);
  const weekKeys = useMemo(() => new Set(weekColumns.map((column) => column.key)), [weekColumns]);
  const jobMap = useMemo(() => new Map(jobs.map((job) => [job.id, job])), [jobs]);
  const jobSearchTextById = useMemo(() => {
    const map = new Map();
    jobs.forEach((job) => {
      map.set(job.id, buildJobSearchText(job));
    });
    return map;
  }, [jobs]);
  const pressAssignmentsByJobId = useMemo(() => {
    const map = new Map();
    assignments.forEach((assignment) => {
      if (assignment.kind !== "press" || !assignment.jobId) return;
      const job = jobMap.get(assignment.jobId);
      if (job && isReleaseJob(job) && assignment.status !== "finished") return;
      const current = map.get(assignment.jobId);
      if (current) {
        current.push(assignment);
      } else {
        map.set(assignment.jobId, [assignment]);
      }
    });
    map.forEach((items) => {
      items.sort((left, right) => {
        if (left.dayKey !== right.dayKey) return left.dayKey.localeCompare(right.dayKey);
        return left.press.localeCompare(right.press);
      });
    });
    return map;
  }, [assignments, jobMap]);
  const normalizedSearch = deferredSearch.trim().toLowerCase();
  const normalizedUnscheduledSearch = deferredUnscheduledSearch.trim().toLowerCase();
  const normalizedLocationSearch = deferredLocationSearch.trim().toLowerCase();

  useEffect(() => {
    if (!weekColumns.length) return;
    setManualScheduleForm((current) =>
      current.dayKey && weekColumns.some((day) => day.key === current.dayKey)
        ? current
        : { ...current, dayKey: weekColumns[0].key, press: PRESS_ORDER.includes(current.press) ? current.press : PRESS_ORDER[0] }
    );
  }, [weekColumns]);

  const filteredJobs = useMemo(() => {
    return jobs.filter((job) => {
      const haystack = jobSearchTextById.get(job.id) || "";
      const matchesSearch = !normalizedSearch || haystack.includes(normalizedSearch);
      return matchesSearch;
    });
  }, [jobSearchTextById, jobs, normalizedSearch]);

  const filteredJobIds = useMemo(() => new Set(filteredJobs.map((job) => job.id)), [filteredJobs]);

  const importedSummary = useMemo(() => {
    const openCount = jobs.filter((job) => job.normalizedStatus === "open").length;
    const closedCount = jobs.filter((job) => job.normalizedStatus === "closed").length;
    return { openCount, closedCount };
  }, [jobs]);

  const userFinishedAssignments = useMemo(
    () =>
      assignments.filter(
        (assignment) =>
          (assignment.kind === "press" || assignment.kind === "finish-only") && assignment.status === "finished"
      ),
    [assignments]
  );

  const finishedMetaByJobId = useMemo(() => {
    const map = new Map();
    userFinishedAssignments.forEach((assignment) => {
      const existing = map.get(assignment.jobId);
      if (!existing || dateSortValue(assignment.finishedAt) > dateSortValue(existing.finishedAt)) {
        map.set(assignment.jobId, {
          finishedAt: assignment.finishedAt || null,
          finishedBy: assignment.finishedBy || "",
          dayKey: assignment.finishedAt ? isoDate(new Date(assignment.finishedAt)) : assignment.dayKey,
          shipDate: assignment.shipDate || (assignment.finishedAt ? isoDate(new Date(assignment.finishedAt)) : assignment.dayKey),
          excludeFromShipping: !!assignment.excludeFromShipping,
        });
      }
    });
    return map;
  }, [userFinishedAssignments]);

  const userFinishedJobIds = useMemo(
    () => new Set(Array.from(finishedMetaByJobId.keys())),
    [finishedMetaByJobId]
  );

  const visibleSchedulerJobIds = useMemo(
    () =>
      new Set(
        filteredJobs
          .filter((job) => !isReleaseJob(job) || userFinishedJobIds.has(job.id))
          .map((job) => job.id)
      ),
    [filteredJobs, userFinishedJobIds]
  );

  const activePressAssignments = useMemo(
    () =>
      assignments.filter(
        (assignment) =>
          assignment.kind === "press" &&
          assignment.status !== "finished" &&
          visibleSchedulerJobIds.has(assignment.jobId)
      ),
    [assignments, visibleSchedulerJobIds]
  );

  const activePressJobIds = useMemo(
    () => new Set(activePressAssignments.map((assignment) => assignment.jobId)),
    [activePressAssignments]
  );

  const activeScheduleLocationsByJobId = useMemo(() => {
    const map = new Map();
    activePressAssignments.forEach((assignment) => {
      const current = map.get(assignment.jobId) || [];
      current.push(assignment);
      map.set(assignment.jobId, current);
    });
    map.forEach((items) => {
      items.sort((left, right) => {
        if (left.dayKey !== right.dayKey) return left.dayKey.localeCompare(right.dayKey);
        return left.press.localeCompare(right.press);
      });
    });
    return map;
  }, [activePressAssignments]);

  const allScheduleLocationsByJobId = useMemo(() => {
    const map = new Map();
    assignments
      .filter((assignment) => assignment.kind === "press" && visibleSchedulerJobIds.has(assignment.jobId))
      .forEach((assignment) => {
        const current = map.get(assignment.jobId) || [];
        current.push(assignment);
        map.set(assignment.jobId, current);
      });
    map.forEach((items) => {
      items.sort((left, right) => {
        if (left.dayKey !== right.dayKey) return left.dayKey.localeCompare(right.dayKey);
        if (left.press !== right.press) return left.press.localeCompare(right.press);
        return dateSortValue(left.finishedAt) - dateSortValue(right.finishedAt);
      });
    });
    return map;
  }, [assignments, visibleSchedulerJobIds]);

  const requestAssigneeOptions = useMemo(
    () => users.map((user) => user.username).sort((left, right) => left.localeCompare(right)),
    [users]
  );

  const unscheduledJobs = useMemo(() => {
    return filteredJobs
      .filter((job) => !isReleaseJob(job))
      .filter((job) => {
        const isScheduled = activePressJobIds.has(job.id);
        if (queueScheduleFilter === "Scheduled") return isScheduled;
        if (queueScheduleFilter === "Not scheduled") return !isScheduled;
        return true;
      })
      .filter((job) => {
        if (queueStatusFilter === "All") return true;
        if (queueStatusFilter === "Open") return job.normalizedStatus === "open";
        return job.normalizedStatus === "closed";
      })
      .filter((job) => {
        if (queuePressFilter === "All") return true;
        return safeText(job.press) === queuePressFilter;
      })
      .filter((job) => {
        if (!queuePriorityFilters.length) return true;
        return queuePriorityFilters.includes(safeText(job.priority));
      })
      .filter((job) => {
        const haystack = jobSearchTextById.get(job.id) || "";
        return !normalizedUnscheduledSearch || haystack.includes(normalizedUnscheduledSearch);
      })
      .sort((left, right) => {
        const leftDate = left.shipByDate ? left.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        const rightDate = right.shipByDate ? right.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        if (leftDate !== rightDate) return leftDate - rightDate;
        return right.estPressTime - left.estPressTime;
      });
  }, [activePressJobIds, filteredJobs, jobSearchTextById, normalizedUnscheduledSearch, queuePressFilter, queuePriorityFilters, queueScheduleFilter, queueStatusFilter]);

  const allUserFinishedJobs = useMemo(() => {
    return jobs
      .filter((job) => userFinishedJobIds.has(job.id))
      .map((job) => ({
        ...job,
        finishMeta: finishedMetaByJobId.get(job.id) || null,
      }))
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [finishedMetaByJobId, jobs, userFinishedJobIds]);

  const doneJobs = useMemo(
    () => allUserFinishedJobs.filter((job) => filteredJobIds.has(job.id)),
    [allUserFinishedJobs, filteredJobIds]
  );

  function compareLaneAssignments(left, right) {
    const leftOrder = Number.isFinite(Number(left.laneOrder)) ? Number(left.laneOrder) : null;
    const rightOrder = Number.isFinite(Number(right.laneOrder)) ? Number(right.laneOrder) : null;
    if (leftOrder != null || rightOrder != null) {
      if (leftOrder == null) return 1;
      if (rightOrder == null) return -1;
      if (leftOrder !== rightOrder) return leftOrder - rightOrder;
    }

    const leftIsManual = left.kind === "manual";
    const rightIsManual = right.kind === "manual";
    if (leftIsManual && rightIsManual) {
      const titleCompare = safeText(left.manualTitle).localeCompare(safeText(right.manualTitle));
      if (titleCompare !== 0) return titleCompare;
      return safeText(left.id).localeCompare(safeText(right.id));
    }
    if (leftIsManual) return -1;
    if (rightIsManual) return 1;

    const leftJob = jobMap.get(left.jobId);
    const rightJob = jobMap.get(right.jobId);
    const estTimeCompare = (rightJob?.estPressTime || 0) - (leftJob?.estPressTime || 0);
    if (estTimeCompare !== 0) return estTimeCompare;
    return safeText(left.id).localeCompare(safeText(right.id));
  }

  function getSortedLaneAssignments(sourceAssignments, dayKey, press, excludedAssignmentId = "") {
    return sourceAssignments
      .filter(
        (assignment) =>
          assignment.dayKey === dayKey && assignment.press === press && assignment.id !== excludedAssignmentId
      )
      .sort(compareLaneAssignments);
  }

  function normalizeLaneAssignments(laneAssignments, dayKey, press) {
    return laneAssignments.map((assignment, index) => ({
      ...assignment,
      dayKey,
      press,
      laneOrder: index,
    }));
  }

  function applyLaneAssignmentUpdates(current, updatedAssignments, addedAssignments = []) {
    const updatedById = new Map(updatedAssignments.map((assignment) => [assignment.id, assignment]));
    const next = current.map((assignment) => updatedById.get(assignment.id) || assignment);
    addedAssignments.forEach((assignment) => {
      const updatedAssignment = updatedById.get(assignment.id) || assignment;
      if (!next.some((item) => item.id === updatedAssignment.id)) {
        next.push(updatedAssignment);
      }
    });
    return next;
  }

  const board = useMemo(() => {
    const map = {};
    weekColumns.forEach((day) => {
      map[day.key] = {};
      PRESS_ORDER.forEach((press) => {
        map[day.key][press] = [];
      });
    });

    assignments.forEach((assignment) => {
      if (!map[assignment.dayKey]) return;
      if (assignment.kind === "manual") {
        map[assignment.dayKey][assignment.press].push({ assignment, job: null });
        return;
      }
      if (assignment.kind !== "press") return;
      if (!visibleSchedulerJobIds.has(assignment.jobId)) return;
      const job = jobMap.get(assignment.jobId);
      if (!job) return;
      map[assignment.dayKey][assignment.press].push({ assignment, job });
    });

    Object.values(map).forEach((pressMap) => {
      Object.values(pressMap).forEach((laneJobs) => {
        laneJobs.sort((left, right) => compareLaneAssignments(left.assignment, right.assignment));
      });
    });

    return map;
  }, [assignments, jobMap, visibleSchedulerJobIds, weekColumns]);

  const openRequests = useMemo(
    () =>
      requests
        .filter((request) => request.status === "open")
        .filter((request) => canUserViewRequest(currentUser, request))
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [currentUser, requests]
  );

  const requestHistory = useMemo(
    () =>
      requests
        .filter((request) => request.status === "done")
        .filter((request) => canUserViewRequest(currentUser, request))
        .filter((request) => !requestHistoryFilterDate || sameDay(request.completedAt, requestHistoryFilterDate))
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [currentUser, requestHistoryFilterDate, requests]
  );

  const pendingRegistrationRequests = useMemo(
    () =>
      registrationRequests
        .filter((request) => request.status === "pending")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [registrationRequests]
  );

  const openPullPaperRequests = useMemo(
    () =>
      pullPaperRequests
        .filter((request) => request.status === "open")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [pullPaperRequests]
  );

  const completedPullPaperRequests = useMemo(
    () =>
      pullPaperRequests
        .filter((request) => request.status === "done")
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [pullPaperRequests]
  );

  const openSuppliesRequests = useMemo(
    () =>
      suppliesRequests
        .filter((request) => request.status === "open")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [suppliesRequests]
  );

  const completedSuppliesRequests = useMemo(
    () =>
      suppliesRequests
        .filter((request) => request.status === "done")
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [suppliesRequests]
  );

  const userNotes = useMemo(
    () =>
      notes
        .filter((note) => comparableUsername(note.ownerUsername) === comparableUsername(currentUser?.username))
        .sort((left, right) => {
          if (left.completed !== right.completed) return left.completed ? 1 : -1;
          return dateSortValue(right.createdAt) - dateSortValue(left.createdAt);
        }),
    [currentUser?.username, notes]
  );

  const shipmentItemsByGroupId = useMemo(() => {
    const map = new Map();
    shipmentGroups.forEach((group) => {
      map.set(group.id, getShipmentItems(group, jobMap, finishedMetaByJobId));
    });
    return map;
  }, [finishedMetaByJobId, jobMap, shipmentGroups]);

  const assignedShipmentJobIds = useMemo(() => {
    const ids = new Set();
    shipmentGroups.forEach((group) => {
      (shipmentItemsByGroupId.get(group.id) || []).forEach((item) => ids.add(item.id));
    });
    return ids;
  }, [shipmentGroups, shipmentItemsByGroupId]);

  const unassignedFinishedJobs = useMemo(() => {
    return allUserFinishedJobs
      .filter((job) => !job.finishMeta?.excludeFromShipping)
      .filter((job) => !assignedShipmentJobIds.has(job.id))
      .filter((job) => isWithinLookbackWindow(job.dateDone || job.finishMeta?.finishedAt, shipQueueWindow))
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [allUserFinishedJobs, assignedShipmentJobIds, shipQueueWindow]);

  useEffect(() => {
    const visibleIds = new Set(unassignedFinishedJobs.map((job) => job.id));
    setSelectedShipQueueJobs((current) => current.filter((jobId) => visibleIds.has(jobId)));
  }, [unassignedFinishedJobs]);

  const dateDoneJobs = useMemo(() => {
    return jobs
      .filter((job) => sameDay(job.dateDone, selectedShipDate))
      .map((job) => ({
        ...job,
        finishMeta: finishedMetaByJobId.get(job.id) || null,
      }))
      .sort((left, right) => {
        const dateDiff = dateSortValue(right.dateDone || right.finishMeta?.finishedAt) - dateSortValue(left.dateDone || left.finishMeta?.finishedAt);
        if (dateDiff !== 0) return dateDiff;
        return parseNumber(right.number) - parseNumber(left.number);
      });
  }, [finishedMetaByJobId, jobs, selectedShipDate]);

  const dateDoneJobNumbers = useMemo(
    () => dateDoneJobs.map((job) => safeText(job.number)).filter(Boolean).join(", "),
    [dateDoneJobs]
  );

  const readyToShipJobs = useMemo(() => {
    return dateDoneJobs
      .filter((job) => !job.finishMeta?.excludeFromShipping)
      .filter((job) => !assignedShipmentJobIds.has(job.id))
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [assignedShipmentJobIds, dateDoneJobs]);

  const shipmentGroupsForDay = useMemo(
    () =>
      shipmentGroups
        .filter((group) => group.shipDate === selectedShipDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [selectedShipDate, shipmentGroups]
  );

  const shipmentHistoryDays = useMemo(() => {
    const grouped = new Map();
    shipmentGroups.forEach((group) => {
      if (shipmentHistoryFilterDate && group.shipDate !== shipmentHistoryFilterDate) return;
      const existing = grouped.get(group.shipDate) || {
        shipDate: group.shipDate,
        groupCount: 0,
        jobCount: 0,
        totalCost: 0,
      };
      const items = shipmentItemsByGroupId.get(group.id) || [];
      grouped.set(group.shipDate, {
        shipDate: group.shipDate,
        groupCount: existing.groupCount + 1,
        jobCount: existing.jobCount + items.length,
        totalCost: existing.totalCost + parseCurrency(group.totalCost),
      });
    });
    return Array.from(grouped.values()).sort((left, right) => right.shipDate.localeCompare(left.shipDate));
  }, [shipmentGroups, shipmentHistoryFilterDate, shipmentItemsByGroupId]);

  const selectedJob = useMemo(
    () => (selectedJobId ? jobMap.get(selectedJobId) || null : null),
    [jobMap, selectedJobId]
  );
  const selectedJobFinishMeta = selectedJob ? finishedMetaByJobId.get(selectedJob.id) : null;

  const jobLocationResults = useMemo(() => {
    if (!normalizedLocationSearch) return [];
    return jobs
      .filter((job) => {
        const haystack = jobSearchTextById.get(job.id) || "";
        return haystack.includes(normalizedLocationSearch);
      })
      .slice(0, 20)
      .map((job) => {
        const locations = pressAssignmentsByJobId.get(job.id) || [];
        return { job, locations };
      });
  }, [jobSearchTextById, jobs, normalizedLocationSearch, pressAssignmentsByJobId]);

  const queuePressOptions = useMemo(
    () => ["All", ...Array.from(new Set(jobs.map((job) => safeText(job.press)).filter(Boolean))).sort()],
    [jobs]
  );

  const queuePriorityOptions = useMemo(
    () => Array.from(new Set(jobs.map((job) => safeText(job.priority)).filter(Boolean))).sort(),
    [jobs]
  );

  const todayDateKey = todayKey();
  const visiblePresses = useMemo(
    () => PRESS_ORDER.filter((press) => visiblePressFilters.includes(press)),
    [visiblePressFilters]
  );
  const visibleScheduleRows = useMemo(
    () =>
      visiblePresses.filter(
        (press) => !hideEmptyPresses || weekColumns.some((day) => (board[day.key]?.[press] || []).length)
      ),
    [board, hideEmptyPresses, visiblePresses, weekColumns]
  );
  const todayScheduledItems = useMemo(
    () =>
      (board[todayDateKey] ? Object.entries(board[todayDateKey]) : [])
        .flatMap(([press, laneJobs]) =>
          laneJobs.map(({ assignment, job }) => ({
            press,
            assignment,
            job,
          }))
        )
        .sort((left, right) => {
          if (!left.job && !right.job) return left.press.localeCompare(right.press);
          if (!left.job) return -1;
          if (!right.job) return 1;
          return left.press.localeCompare(right.press) || right.job.estPressTime - left.job.estPressTime;
        }),
    [board, todayDateKey]
  );
  const jobsDueToday = useMemo(
    () =>
      jobs
        .filter((job) => !isReleaseJob(job))
        .filter((job) => job.shipByDate && isoDate(job.shipByDate) === todayDateKey && !userFinishedJobIds.has(job.id))
        .sort((left, right) => left.estPressTime - right.estPressTime),
    [jobs, todayDateKey, userFinishedJobIds]
  );
  const finishedTodayJobs = useMemo(
    () => allUserFinishedJobs.filter((job) => sameDay(job.finishMeta?.finishedAt, todayDateKey)),
    [allUserFinishedJobs, todayDateKey]
  );
  const shipGroupsToday = useMemo(
    () => shipmentGroups.filter((group) => group.shipDate === todayDateKey),
    [shipmentGroups, todayDateKey]
  );
  const shipmentEmailSentToday = useMemo(
    () => shipmentEmailLogs.some((log) => log.shipDate === todayDateKey),
    [shipmentEmailLogs, todayDateKey]
  );
  const todayDashboardStats = useMemo(
    () => [
      { label: "Due today", value: jobsDueToday.length },
      { label: "Scheduled today", value: todayScheduledItems.filter((item) => item.assignment.kind === "press").length },
      { label: "Finished today", value: finishedTodayJobs.length },
      { label: "Open requests", value: openRequests.length },
      { label: "Pull paper", value: openPullPaperRequests.length },
      { label: "Supplies", value: openSuppliesRequests.length },
    ],
    [finishedTodayJobs.length, jobsDueToday.length, openPullPaperRequests.length, openRequests.length, openSuppliesRequests.length, todayScheduledItems]
  );
  const activityActors = useMemo(
    () => ["All", ...Array.from(new Set(activityLog.map((entry) => safeText(entry.actor)).filter(Boolean))).sort()],
    [activityLog]
  );
  const activityScopes = useMemo(
    () => ["All", ...Array.from(new Set(activityLog.map((entry) => safeText(entry.scope)).filter(Boolean))).sort()],
    [activityLog]
  );
  const filteredActivityLog = useMemo(
    () =>
      activityLog
        .filter((entry) => activityFilters.actor === "All" || safeText(entry.actor) === activityFilters.actor)
        .filter((entry) => activityFilters.scope === "All" || safeText(entry.scope) === activityFilters.scope)
        .filter((entry) => !activityFilters.date || sameDay(entry.createdAt, activityFilters.date))
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [activityFilters.actor, activityFilters.date, activityFilters.scope, activityLog]
  );

  const summary = useMemo(() => {
    return {
      txtOpen: importedSummary.openCount,
      txtClosed: importedSummary.closedCount,
      markedFinished: allUserFinishedJobs.length,
      openRequests: openRequests.length,
      scheduledJobs: activePressAssignments.length,
      shipGroupsOnDate: shipmentGroupsForDay.length,
    };
  }, [activePressAssignments.length, allUserFinishedJobs.length, importedSummary.closedCount, importedSummary.openCount, openRequests.length, shipmentGroupsForDay.length]);

  const shipmentEmailsForSelectedDate = useMemo(
    () =>
      shipmentEmailLogs
        .filter((log) => log.shipDate === selectedShipDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [selectedShipDate, shipmentEmailLogs]
  );

  const shipmentEmailDraft = useMemo(() => {
    const groups = shipmentGroupsForDay;
    const totalCost = groups.reduce((sum, group) => sum + parseCurrency(group.totalCost), 0);
    const totalBill = groups.reduce((sum, group) => sum + parseCurrency(group.billAmount), 0);
    const methods = Array.from(new Set(groups.map((group) => group.method).filter(Boolean)));
    const lines = [`Daily shipment summary for ${selectedShipDate}`, ""];
    let jobCount = 0;

    groups.forEach((group) => {
      const items = shipmentItemsByGroupId.get(group.id) || [];
      jobCount += items.length;
      lines.push(`${group.label} | ${group.method}`);
      lines.push(`Skids / cartons: ${group.packageCount || 0} ${group.packageType || "Skids"}`);
      lines.push(`Our cost: ${formatCurrency(group.totalCost)}`);
      lines.push(`Bill: ${formatCurrency(group.billAmount)}`);
      items.forEach((item) => {
        lines.push(`- ${item.customerName} ${item.number}: ${item.generalDescr}`);
      });
      if (group.notes) lines.push(`Notes: ${group.notes}`);
      lines.push("");
    });

    if (!groups.length) {
      lines.push("No shipment groups have been created for this date yet.");
      lines.push("");
    }

    lines.push(`Total cost: ${formatCurrency(totalCost)}`);
    lines.push(`Total bill: ${formatCurrency(totalBill)}`);

    return {
      subject: `Daily shipments for ${selectedShipDate}`,
      body: lines.join("\n"),
      groupCount: groups.length,
      jobCount,
      totalCost,
      totalBill,
      methods,
    };
  }, [selectedShipDate, shipmentGroupsForDay, shipmentItemsByGroupId]);

  const shipmentEmailHistory = useMemo(
    () =>
      shipmentEmailLogs
        .filter((log) => !shipmentEmailHistoryFilterDate || log.shipDate === shipmentEmailHistoryFilterDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [shipmentEmailHistoryFilterDate, shipmentEmailLogs]
  );

  useEffect(() => {
    const validIds = new Set(readyToShipJobs.map((job) => job.id));
    setSelectedShipmentJobs((current) => current.filter((jobId) => validIds.has(jobId)));
  }, [readyToShipJobs]);

  useEffect(() => {
    const validIds = new Set(unassignedFinishedJobs.map((job) => job.id));
    setSelectedShipQueueJobs((current) => current.filter((jobId) => validIds.has(jobId)));
  }, [unassignedFinishedJobs]);

  function selectJob(jobId, shouldScroll = false) {
    setSelectedJobId(jobId);
    if (!shouldScroll) return;
    requestAnimationFrame(() => {
      jobDetailsRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  function importText(text) {
    const parsed = parseLabelTraxxText(text);
    if (!parsed.length) return;
    const parsedById = new Map(parsed.map((job) => [job.id, job]));
    setJobs((current) => {
      const currentById = new Map(current.map((job) => [job.id, job]));
      return parsed.map((job) => {
        const previous = currentById.get(job.id);
        return previous
          ? {
              ...job,
              holdActive: !!previous.holdActive,
              holdNote: safeText(previous.holdNote),
            }
          : {
              ...job,
              holdActive: false,
              holdNote: "",
            };
      });
    });
    setAssignments((current) => {
      const scopedAssignments = current.filter(
        (assignment) => assignment.kind === "manual" || parsedById.has(assignment.jobId)
      );
      const nextAssignments = scopedAssignments.filter(
        (assignment) => !(assignment.kind === "finish-only" && assignment.importedFinish)
      );

      parsed.forEach((job) => {
        if (!job.dateDone) return;
        const hasFinishedAssignment = nextAssignments.some(
          (assignment) =>
            assignment.jobId === job.id &&
            (assignment.kind === "press" || assignment.kind === "finish-only") &&
            assignment.status === "finished"
        );
        if (hasFinishedAssignment) return;
        const doneDayKey = isoDate(job.dateDone);
        const importedFinishedAt = localMiddayIso(job.dateDone);
        nextAssignments.push({
          id: makeId("asg"),
          jobId: job.id,
          dayKey: doneDayKey,
          press: job.press && PRESS_ORDER.includes(job.press) ? job.press : "Rewind",
          kind: "finish-only",
          status: "finished",
          createdAt: importedFinishedAt,
          finishedAt: importedFinishedAt,
          finishedBy: "Import",
          shipDate: doneDayKey,
          importedFinish: true,
          excludeFromShipping: false,
        });
      });

      return nextAssignments;
    });
    setShipmentGroups((current) =>
      current.map((group) => {
        const nextJobMap = new Map(parsed.map((job) => [job.id, job]));
        const jobItems = getShipmentItems(group, nextJobMap, finishedMetaByJobId).filter((item) =>
          parsed.some((job) => job.id === item.id)
        );
        return { ...group, jobItems };
      })
    );
    const firstDate = parsed.find((job) => job.shipByDate)?.shipByDate;
    if (firstDate) setWeekStart(startOfWeek(firstDate));
    recordActivity("Imported TXT schedule", "Scheduler", `${parsed.length} jobs were loaded into the scheduler.`);
  }

  function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => importText(String(reader.result || ""));
    reader.readAsText(file);
  }

  async function handleScheduleWorkbookUpload(event) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file || !userCanManageUsers) return;

    try {
      const buffer = await file.arrayBuffer();
      const parsed = importWeeklyScheduleWorkbook(buffer, jobs);
      if (!parsed.sheetCount) {
        window.alert("No weekly schedule tabs were found in that workbook.");
        return;
      }

      const importedDays = new Set(parsed.importedDayKeys);
      setJobs(parsed.jobs);
      setAssignments((current) => [
        ...current.filter((assignment) => assignment.status === "finished" || !importedDays.has(assignment.dayKey)),
        ...parsed.assignments,
      ]);
      if (parsed.weekStart) setWeekStart(parsed.weekStart);

      recordActivity(
        "Imported Excel schedule",
        "Scheduler",
        `${file.name} added ${parsed.assignments.length} schedule item(s) from ${parsed.sheetCount} weekly tab(s).`,
        {
          fileName: file.name,
          sheetCount: parsed.sheetCount,
          assignmentCount: parsed.assignments.length,
        }
      );
    } catch (error) {
      console.error("Failed to import Excel schedule.", error);
      window.alert("That workbook could not be imported. Please check the weekly sheet layout and try again.");
    }
  }

  function handleScheduleDrop(event, dayKey, press) {
    event.preventDefault();
    if (!userCanMoveJobs) return;
    const payload = parseDragPayload(event.dataTransfer.getData("application/json"));
    if (payload?.type === "queue" && payload.jobId) {
      addAssignment(payload.jobId, dayKey, press);
      setPickedUpItem(null);
      return;
    }
    if (payload?.type === "scheduled" && payload.assignmentId) {
      moveAssignment(payload.assignmentId, dayKey, press);
      setPickedUpItem(null);
      return;
    }
    const fallbackJobId = event.dataTransfer.getData("text/plain");
    if (fallbackJobId) {
      addAssignment(fallbackJobId, dayKey, press);
      setPickedUpItem(null);
    }
  }

  function handleScheduleCardDrop(event, dayKey, press, targetAssignmentId) {
    event.preventDefault();
    event.stopPropagation();
    if (!userCanMoveJobs) return;
    const payload = parseDragPayload(event.dataTransfer.getData("application/json"));
    if (payload?.type === "queue" && payload.jobId) {
      addAssignment(payload.jobId, dayKey, press, targetAssignmentId);
      setPickedUpItem(null);
      return;
    }
    if (payload?.type === "scheduled" && payload.assignmentId) {
      if (payload.assignmentId === targetAssignmentId) return;
      moveAssignment(payload.assignmentId, dayKey, press, targetAssignmentId);
      setPickedUpItem(null);
      return;
    }
    const fallbackJobId = event.dataTransfer.getData("text/plain");
    if (fallbackJobId) {
      addAssignment(fallbackJobId, dayKey, press, targetAssignmentId);
      setPickedUpItem(null);
    }
  }

  function pickUpQueueJob(jobId) {
    if (!userCanMoveJobs) return;
    const job = jobMap.get(jobId);
    if (!job) return;
    setPickedUpItem({
      type: "queue",
      jobId,
      label: `${job.customerName} ${job.number}`.trim(),
    });
  }

  function pickUpScheduledAssignment(assignmentId) {
    if (!userCanMoveJobs) return;
    const assignment = assignments.find((item) => item.id === assignmentId);
    if (!assignment) return;
    const job = assignment.jobId ? jobMap.get(assignment.jobId) : null;
    setPickedUpItem({
      type: "scheduled",
      assignmentId,
      jobId: assignment.jobId,
      label: assignment.kind === "manual" ? assignment.manualTitle : `${job?.customerName || ""} ${job?.number || ""}`.trim(),
    });
  }

  function placePickedUpItem(dayKey, press) {
    if (!userCanMoveJobs || !pickedUpItem) return;
    if (pickedUpItem.type === "queue" && pickedUpItem.jobId) {
      addAssignment(pickedUpItem.jobId, dayKey, press);
      setPickedUpItem(null);
      return;
    }
    if (pickedUpItem.type === "scheduled" && pickedUpItem.assignmentId) {
      moveAssignment(pickedUpItem.assignmentId, dayKey, press);
      setPickedUpItem(null);
    }
  }

  function toggleQueuePriority(priority) {
    const value = safeText(priority);
    if (!value) return;
    setQueuePriorityFilters((current) =>
      current.includes(value) ? current.filter((item) => item !== value) : [...current, value]
    );
  }

  function toggleVisiblePressFilter(press) {
    setVisiblePressFilters((current) =>
      current.includes(press) ? current.filter((item) => item !== press) : [...current, press]
    );
  }

  function addAssignment(jobId, dayKey, press, beforeAssignmentId = "") {
    if (!userCanMoveJobs) return;
    const job = jobMap.get(jobId);
    if (job && isReleaseJob(job)) return;
    const exists = assignments.some(
      (assignment) =>
        assignment.jobId === jobId &&
        assignment.dayKey === dayKey &&
        assignment.press === press &&
        assignment.kind === "press" &&
        assignment.status !== "finished"
    );
    if (exists) return;
    const newAssignment = {
      id: makeId("asg"),
      jobId,
      dayKey,
      press,
      kind: "press",
      status: "scheduled",
      createdAt: new Date().toISOString(),
      finishedAt: null,
      finishedBy: "",
      laneOrder: null,
    };
    setAssignments((current) => {
      const laneAssignments = getSortedLaneAssignments(current, dayKey, press);
      const targetIndex = beforeAssignmentId
        ? laneAssignments.findIndex((assignment) => assignment.id === beforeAssignmentId)
        : -1;
      const insertAt = targetIndex >= 0 ? targetIndex : laneAssignments.length;
      const nextLane = normalizeLaneAssignments(
        [
          ...laneAssignments.slice(0, insertAt),
          newAssignment,
          ...laneAssignments.slice(insertAt),
        ],
        dayKey,
        press
      );
      return applyLaneAssignmentUpdates(current, nextLane, [newAssignment]);
    });
    recordActivity(
      "Scheduled job",
      "Scheduler",
      `${job ? `${job.customerName} ${job.number}` : "A job"} was placed on Press ${press} for ${dayKey}.`,
      { jobId, dayKey, press }
    );
  }

  function addManualScheduleEntry(event) {
    event.preventDefault();
    if (!userCanMoveJobs) return;
    const title = safeText(manualScheduleForm.title);
    const dayKey = safeText(manualScheduleForm.dayKey);
    const press = safeText(manualScheduleForm.press);
    if (!title || !dayKey || !PRESS_ORDER.includes(press)) return;

    const newAssignment = {
      id: makeId("manual"),
      jobId: "",
      manualTitle: title,
      dayKey,
      press,
      kind: "manual",
      status: "scheduled",
      createdAt: new Date().toISOString(),
      finishedAt: null,
      finishedBy: "",
      laneOrder: null,
    };

    setAssignments((current) => {
      const laneAssignments = getSortedLaneAssignments(current, dayKey, press);
      const nextLane = normalizeLaneAssignments([...laneAssignments, newAssignment], dayKey, press);
      return applyLaneAssignmentUpdates(current, nextLane, [newAssignment]);
    });
    recordActivity(
      "Added manual block",
      "Scheduler",
      `${title} was added to Press ${press} on ${dayKey}.`,
      { title, dayKey, press }
    );
    setManualScheduleForm((current) => ({ ...current, title: "" }));
  }

  function moveAssignment(assignmentId, dayKey, press, beforeAssignmentId = "") {
    if (!userCanMoveJobs) return;
    const assignmentToMove = assignments.find((assignment) => assignment.id === assignmentId);
    if (!assignmentToMove) return;
    const movingWithinLane = assignmentToMove.dayKey === dayKey && assignmentToMove.press === press;
    if (assignmentToMove.kind === "press" && assignmentToMove.status === "finished" && !movingWithinLane) return;
    const duplicate = assignments.some(
      (assignment) =>
        assignment.id !== assignmentId &&
        assignment.dayKey === dayKey &&
        assignment.press === press &&
        (
          (assignmentToMove.kind === "manual" &&
            assignment.kind === "manual" &&
            comparableUsername(assignment.manualTitle) === comparableUsername(assignmentToMove.manualTitle)) ||
          (assignmentToMove.kind === "press" &&
            assignment.jobId === assignmentToMove.jobId &&
            assignment.kind === "press" &&
            assignment.status !== "finished")
        )
    );
    if (duplicate && !movingWithinLane) return;

    setAssignments((current) => {
      const currentAssignment = current.find((assignment) => assignment.id === assignmentId);
      if (!currentAssignment) return current;
      const isSameLane = currentAssignment.dayKey === dayKey && currentAssignment.press === press;
      const sourceDayKey = currentAssignment.dayKey;
      const sourcePress = currentAssignment.press;
      const sourceLane = getSortedLaneAssignments(current, sourceDayKey, sourcePress, assignmentId);
      const destinationBase = isSameLane ? sourceLane : getSortedLaneAssignments(current, dayKey, press);
      const targetIndex = beforeAssignmentId
        ? destinationBase.findIndex((assignment) => assignment.id === beforeAssignmentId)
        : -1;
      const insertAt = targetIndex >= 0 ? targetIndex : destinationBase.length;
      const movedAssignment = {
        ...currentAssignment,
        dayKey,
        press,
      };
      const destinationLane = normalizeLaneAssignments(
        [
          ...destinationBase.slice(0, insertAt),
          movedAssignment,
          ...destinationBase.slice(insertAt),
        ],
        dayKey,
        press
      );
      if (isSameLane) {
        return applyLaneAssignmentUpdates(current, destinationLane);
      }
      const updatedSourceLane = normalizeLaneAssignments(sourceLane, sourceDayKey, sourcePress);
      return applyLaneAssignmentUpdates(current, [...destinationLane, ...updatedSourceLane]);
    });
    const job = assignmentToMove.jobId ? jobMap.get(assignmentToMove.jobId) : null;
    const title = assignmentToMove.kind === "manual" ? assignmentToMove.manualTitle : `${job?.customerName || ""} ${job?.number || ""}`.trim();
    recordActivity(
      movingWithinLane ? "Reordered schedule item" : "Moved schedule item",
      "Scheduler",
      movingWithinLane
        ? `${title || "Schedule item"} was reordered on Press ${press} for ${dayKey}.`
        : `${title || "Schedule item"} moved to Press ${press} on ${dayKey}.`,
      {
        assignmentId,
        fromDay: assignmentToMove.dayKey,
        fromPress: assignmentToMove.press,
        toDay: dayKey,
        toPress: press,
        beforeAssignmentId,
      }
    );
  }

  function moveAssignmentByStep(assignmentId, direction) {
    if (!userCanMoveJobs) return;
    const assignmentToMove = assignments.find((assignment) => assignment.id === assignmentId);
    if (!assignmentToMove) return;
    const laneAssignments = getSortedLaneAssignments(assignments, assignmentToMove.dayKey, assignmentToMove.press);
    const currentIndex = laneAssignments.findIndex((assignment) => assignment.id === assignmentId);
    const targetIndex = currentIndex + direction;
    if (currentIndex < 0 || targetIndex < 0 || targetIndex >= laneAssignments.length) return;

    setAssignments((current) => {
      const currentAssignment = current.find((assignment) => assignment.id === assignmentId);
      if (!currentAssignment) return current;
      const currentLane = getSortedLaneAssignments(current, currentAssignment.dayKey, currentAssignment.press);
      const liveIndex = currentLane.findIndex((assignment) => assignment.id === assignmentId);
      const liveTargetIndex = liveIndex + direction;
      if (liveIndex < 0 || liveTargetIndex < 0 || liveTargetIndex >= currentLane.length) return current;
      const reorderedLane = [...currentLane];
      [reorderedLane[liveIndex], reorderedLane[liveTargetIndex]] = [reorderedLane[liveTargetIndex], reorderedLane[liveIndex]];
      return applyLaneAssignmentUpdates(
        current,
        normalizeLaneAssignments(reorderedLane, currentAssignment.dayKey, currentAssignment.press)
      );
    });

    const job = assignmentToMove.jobId ? jobMap.get(assignmentToMove.jobId) : null;
    const title = assignmentToMove.kind === "manual" ? assignmentToMove.manualTitle : `${job?.customerName || ""} ${job?.number || ""}`.trim();
    recordActivity(
      "Reordered schedule item",
      "Scheduler",
      `${title || "Schedule item"} moved ${direction < 0 ? "up" : "down"} on Press ${assignmentToMove.press} for ${assignmentToMove.dayKey}.`,
      { assignmentId, direction, dayKey: assignmentToMove.dayKey, press: assignmentToMove.press }
    );
  }

  function duplicateAssignmentToNextDay(assignmentId) {
    if (!userCanMoveJobs) return;
    const assignmentToCopy = assignments.find((assignment) => assignment.id === assignmentId);
    if (!assignmentToCopy) return;
    if (assignmentToCopy.kind === "press" && assignmentToCopy.status === "finished") return;

    const currentIndex = weekColumns.findIndex((day) => day.key === assignmentToCopy.dayKey);
    if (currentIndex < 0 || currentIndex >= weekColumns.length - 1) return;

    const nextDayKey = weekColumns[currentIndex + 1].key;
    const exists = assignments.some(
      (assignment) =>
        assignment.id !== assignmentId &&
        assignment.dayKey === nextDayKey &&
        assignment.press === assignmentToCopy.press &&
        (
          (assignmentToCopy.kind === "manual" &&
            assignment.kind === "manual" &&
            comparableUsername(assignment.manualTitle) === comparableUsername(assignmentToCopy.manualTitle)) ||
          (assignmentToCopy.kind === "press" &&
            assignment.jobId === assignmentToCopy.jobId &&
            assignment.kind === "press" &&
            assignment.status !== "finished")
        )
    );
    if (exists) return;

    const nextAssignment = {
      ...assignmentToCopy,
      id: makeId("asg"),
      dayKey: nextDayKey,
      createdAt: new Date().toISOString(),
      finishedAt: null,
      finishedBy: "",
      status: "scheduled",
      laneOrder: null,
    };
    setAssignments((current) => {
      const laneAssignments = getSortedLaneAssignments(current, nextDayKey, assignmentToCopy.press);
      const nextLane = normalizeLaneAssignments([...laneAssignments, nextAssignment], nextDayKey, assignmentToCopy.press);
      return applyLaneAssignmentUpdates(current, nextLane, [nextAssignment]);
    });
    const job = assignmentToCopy.jobId ? jobMap.get(assignmentToCopy.jobId) : null;
    const title = assignmentToCopy.kind === "manual" ? assignmentToCopy.manualTitle : `${job?.customerName || ""} ${job?.number || ""}`.trim();
    recordActivity(
      "Duplicated schedule item",
      "Scheduler",
      `${title || "Schedule item"} was copied to ${nextDayKey} on Press ${assignmentToCopy.press}.`,
      { assignmentId, nextDayKey, press: assignmentToCopy.press }
    );
  }

  function removeAssignment(assignmentId) {
    if (!userCanMoveJobs) return;
    const assignment = assignments.find((item) => item.id === assignmentId);
    if (!assignment) return;
    setAssignments((current) => {
      const next = current.filter((item) => item.id !== assignmentId);
      const updatedLane = normalizeLaneAssignments(
        getSortedLaneAssignments(next, assignment.dayKey, assignment.press),
        assignment.dayKey,
        assignment.press
      );
      return applyLaneAssignmentUpdates(next, updatedLane);
    });
    const job = assignment.jobId ? jobMap.get(assignment.jobId) : null;
    const title = assignment.kind === "manual" ? assignment.manualTitle : `${job?.customerName || ""} ${job?.number || ""}`.trim();
    recordActivity(
      "Removed schedule item",
      "Scheduler",
      `${title || "Schedule item"} was removed from Press ${assignment.press} on ${assignment.dayKey}.`,
      { assignmentId, dayKey: assignment.dayKey, press: assignment.press }
    );
  }

  function getPressOperator(dayKey, press) {
    return pressOperators[pressOperatorKey(dayKey, press)] || "";
  }

  function updatePressOperator(dayKey, press, value) {
    if (!userCanEdit) return;
    const key = pressOperatorKey(dayKey, press);
    const nextValue = safeText(value);
    setPressOperators((current) => {
      if (nextValue) return { ...current, [key]: nextValue };
      if (!(key in current)) return current;
      const next = { ...current };
      delete next[key];
      return next;
    });
  }

  function commitPressOperator(dayKey, press, previousValue, value) {
    if (!currentUser || !userCanEdit) return;
    const previous = safeText(previousValue);
    const nextValue = safeText(value);
    if (previous === nextValue) return;
    recordActivity(
      nextValue ? "Updated press operator" : "Cleared press operator",
      "Scheduler",
      nextValue
        ? `${nextValue} is assigned to Press ${press} on ${dayKey}.`
        : `The operator was cleared from Press ${press} on ${dayKey}.`,
      { dayKey, press, operator: nextValue }
    );
  }

  function getPressDuty(dayKey, press) {
    return pressDuties[pressOperatorKey(dayKey, press)] || "";
  }

  function updatePressDuty(dayKey, press, value) {
    if (!userCanEdit) return;
    const key = pressOperatorKey(dayKey, press);
    const nextValue = safeText(value);
    setPressDuties((current) => {
      if (nextValue) return { ...current, [key]: nextValue };
      if (!(key in current)) return current;
      const next = { ...current };
      delete next[key];
      return next;
    });
  }

  function commitPressDuty(dayKey, press, previousValue, value) {
    if (!currentUser || !userCanEdit) return;
    const previous = safeText(previousValue);
    const nextValue = safeText(value);
    if (previous === nextValue) return;
    recordActivity(
      nextValue ? "Updated press duty" : "Cleared press duty",
      "Scheduler",
      nextValue
        ? `${formatPressLabel(press)} on ${dayKey} is set to ${nextValue}.`
        : `The duty was cleared from ${formatPressLabel(press)} on ${dayKey}.`,
      { dayKey, press, duty: nextValue }
    );
  }

  function finishJob(jobId) {
    if (!currentUser || !userCanEdit) return;
    const finishedAt = new Date().toISOString();
    const defaultShipDate = isoDate(new Date(finishedAt));
    const finishedBy = currentUser.username;
    const job = jobMap.get(jobId);
    const fallbackDayKey = job?.shipByDate ? isoDate(job.shipByDate) : weekColumns[0]?.key || todayKey();
    const fallbackPress = job?.press && PRESS_ORDER.includes(job.press) ? job.press : "Rewind";

    setAssignments((current) => {
      let foundPress = false;
      let next = current.map((assignment) => {
        if (assignment.jobId !== jobId) return assignment;
        foundPress = true;
        return {
          ...assignment,
          status: "finished",
          finishedAt,
          finishedBy,
          shipDate: assignment.shipDate || defaultShipDate,
        };
      });

      if (!foundPress) {
        next = [
          ...next,
          {
            id: makeId("asg"),
            jobId,
            dayKey: fallbackDayKey,
            press: fallbackPress,
            kind: "finish-only",
            status: "finished",
            createdAt: finishedAt,
            finishedAt,
            finishedBy,
            shipDate: defaultShipDate,
          },
        ];
      }

      return next;
    });
    if (job) {
      recordActivity("Marked job done", "Scheduler", `${job.customerName} ${job.number} was marked done.`, {
        jobId,
        shipDate: defaultShipDate,
      });
    }
  }

  function undoFinishJob(jobId) {
    if (!currentUser || !userCanEdit) return;
    const job = jobMap.get(jobId);
    const finishedAssignments = assignments.filter(
      (assignment) =>
        assignment.jobId === jobId &&
        (assignment.kind === "press" || assignment.kind === "finish-only") &&
        assignment.status === "finished"
    );
    if (!finishedAssignments.length) return;

    setAssignments((current) =>
      current
        .filter(
          (assignment) =>
            !(
              assignment.jobId === jobId &&
              assignment.kind === "finish-only" &&
              assignment.status === "finished"
            )
        )
        .map((assignment) =>
          assignment.jobId === jobId && assignment.kind === "press" && assignment.status === "finished"
            ? {
                ...assignment,
                status: "scheduled",
                finishedAt: null,
                finishedBy: "",
                excludeFromShipping: false,
              }
            : assignment
        )
    );
    setSelectedShipmentJobs((current) => current.filter((value) => value !== jobId));
    setSelectedShipQueueJobs((current) => current.filter((value) => value !== jobId));
    if (job) {
      recordActivity("Reopened finished job", "Scheduler", `${job.customerName} ${job.number} was marked back to scheduled.`, {
        jobId,
      });
    }
  }

  function clearBoard() {
    if (!userCanMoveJobs) return;
    if (!assignments.length) return;
    setAssignments([]);
    recordActivity("Cleared schedule board", "Scheduler", `${assignments.length} scheduled items were cleared from the board.`);
  }

  function exportSchedule() {
    const rows = [
      [
        "Day",
        "Press",
        "Number",
        "CustomerName",
        "Description",
        "Priority",
        "EST Time",
        "Quantity",
        "Footage",
        "Stock",
        "Ship Date",
        "Imported Status",
        "Marked Finished By",
        "Marked Finished At",
      ],
    ];

    weekColumns.forEach((day) => {
      PRESS_ORDER.forEach((press) => {
        (board[day.key]?.[press] || []).forEach(({ job }) => {
          const finishMeta = finishedMetaByJobId.get(job.id);
          rows.push([
            day.label,
            press,
            job.number,
            job.customerName,
            job.generalDescr,
            job.priority,
            job.estPressTime,
            job.ticQuantity,
            job.estFootage,
            job.stockDisplay,
            formatDate(job.shipByDate),
            job.ticketStatus,
            finishMeta?.finishedBy || "",
            finishMeta?.finishedAt || "",
          ]);
        });
      });
    });

    doneJobs.forEach((job) => {
      rows.push([
        "Ship Ready",
        "History",
        job.number,
        job.customerName,
        job.generalDescr,
        job.priority,
        job.estPressTime,
        job.ticQuantity,
        job.estFootage,
        job.stockDisplay,
        formatDate(job.shipByDate),
        job.ticketStatus,
        job.finishMeta?.finishedBy || "",
        job.finishMeta?.finishedAt || "",
      ]);
    });

    const csv = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
    downloadFile(`schedule-${isoDate(weekStart)}.csv`, csv, "text/csv;charset=utf-8");
  }

  async function exportScheduleWorkbook() {
    if (!userCanManageUsers) return;

    const lanesByDay = {};
    const finishedRowsByDay = {};

    weekColumns.forEach((day) => {
      lanesByDay[day.key] = {};
      PRESS_ORDER.forEach((press) => {
        lanesByDay[day.key][press] = (board[day.key]?.[press] || [])
          .filter(({ assignment }) => assignment.status !== "finished")
          .map(({ assignment, job }) =>
            assignment.kind === "manual"
              ? {
                  kind: "manual",
                  title: assignment.manualTitle,
                }
              : {
                  kind: "job",
                  label: `${job.customerName} ${job.number}`.trim(),
                  estTime: job.estPressTime || "",
                  quantity: job.ticQuantity || "",
                  footage: job.estFootage || "",
                  stock: job.stockDisplay || "",
                  shipDate: job.shipByDate ? formatDate(job.shipByDate) : "",
                }
          );
      });

      const groupRows = [];
      const shippedJobIds = new Set();
      shipmentGroups
        .filter((group) => group.shipDate === day.key)
        .forEach((group) => {
          const packageType = safeText(group.packageType).toLowerCase();
          const items = shipmentItemsByGroupId.get(group.id) || [];
          items.forEach((item) => {
            shippedJobIds.add(item.id);
            const liveJob = jobMap.get(item.id);
            groupRows.push({
              label: `${item.customerName} ${item.number}`.trim(),
              totalCartons: packageType.includes("carton") ? group.packageCount || "" : "",
              quantity: liveJob?.ticQuantity || "",
              skids: packageType.includes("skid") ? group.packageCount || "" : "",
              method: group.method,
              cost: group.totalCost ? formatCurrency(parseCurrency(group.totalCost)) : "",
            });
          });
        });

      allUserFinishedJobs
        .filter((job) => effectiveFinishedShipDate(job.finishMeta) === day.key)
        .forEach((job) => {
          if (shippedJobIds.has(job.id)) return;
          groupRows.push({
            label: `${job.customerName} ${job.number}`.trim(),
            totalCartons: "",
            quantity: job.ticQuantity || "",
            skids: "",
            method: "",
            cost: "",
          });
        });

      finishedRowsByDay[day.key] = groupRows;
    });

    const workbookBuffer = await exportWeeklyScheduleWorkbook({
      weekStart,
      weekColumns,
      lanesByDay,
      finishedRowsByDay,
    });

    downloadFile(
      `schedule-${isoDate(weekStart)}.xlsx`,
      workbookBuffer,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    recordActivity("Exported Excel schedule", "Scheduler", `The week of ${isoDate(weekStart)} was exported to Excel.`);
  }

  async function handleDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setRequestDraftAttachments((current) => [...current, ...nextAttachments]);
      recordActivity("Added request draft files", "Requests", `${nextAttachments.length} file(s) added to a new request draft.`);
    }
    event.target.value = "";
  }

  function removeDraftAttachment(attachmentId) {
    const attachment = requestDraftAttachments.find((item) => item.id === attachmentId);
    setRequestDraftAttachments((current) => current.filter((item) => item.id !== attachmentId));
    deleteStoredAttachments(attachment ? [attachment] : []);
  }

  async function handleSuppliesDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setSuppliesDraftAttachments((current) => [...current, ...nextAttachments]);
      recordActivity("Added supplies draft files", "Supplies", `${nextAttachments.length} file(s) added to a supplies request draft.`);
    }
    event.target.value = "";
  }

  function removeSuppliesDraftAttachment(attachmentId) {
    const attachment = suppliesDraftAttachments.find((item) => item.id === attachmentId);
    setSuppliesDraftAttachments((current) => current.filter((item) => item.id !== attachmentId));
    deleteStoredAttachments(attachment ? [attachment] : []);
  }

  async function handleShipmentDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setShipmentDraftAttachments((current) => [...current, ...nextAttachments]);
      recordActivity("Added shipment draft files", "Shipping", `${nextAttachments.length} file(s) added to a shipment draft.`);
    }
    event.target.value = "";
  }

  function removeShipmentDraftAttachment(attachmentId) {
    const attachment = shipmentDraftAttachments.find((item) => item.id === attachmentId);
    setShipmentDraftAttachments((current) => current.filter((item) => item.id !== attachmentId));
    deleteStoredAttachments(attachment ? [attachment] : []);
  }

  function clearRequestDraft() {
    deleteStoredAttachments(requestDraftAttachments);
    setRequestForm(EMPTY_REQUEST_FORM);
    setRequestDraftAttachments([]);
  }

  function clearSuppliesDraft() {
    deleteStoredAttachments(suppliesDraftAttachments);
    setSuppliesForm(EMPTY_SUPPLIES_FORM);
    setSuppliesDraftAttachments([]);
  }

  function clearShipmentDraft() {
    deleteStoredAttachments(shipmentDraftAttachments);
    setShipmentForm({ ...EMPTY_SHIPMENT_FORM, shipDate: selectedShipDate });
    setShipmentDraftAttachments([]);
  }

  function submitRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!requestForm.jobNumber || !requestForm.customer || !requestForm.requestorName || !requestForm.description) {
      return;
    }

    const assignedToAccount = safeText(requestForm.assignedToAccount);
    const requestType = safeText(requestForm.requestType) || "General Request";

    setRequests((current) => [
      {
        id: makeId("req"),
        ...requestForm,
        assignedToAccount,
        requestType,
        attachments: requestDraftAttachments,
        createdAt: new Date().toISOString(),
        createdByAccount: currentUser.username,
        completedAt: null,
        completedByAccount: "",
        status: "open",
      },
      ...current,
    ]);
    recordActivity(
      "Created request",
      "Requests",
      `${requestForm.customer} ${requestForm.jobNumber} was added to open requests${assignedToAccount ? ` for ${assignedToAccount}` : ""}.`,
      { requestType, assignedToAccount }
    );
    setRequestForm(EMPTY_REQUEST_FORM);
    setRequestDraftAttachments([]);
    setActiveTab("Open Requests");
  }

  function submitPullPaperRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!pullPaperForm.details.trim()) return;

    setPullPaperRequests((current) => [
      {
        id: makeId("paper"),
        details: pullPaperForm.details.trim(),
        target: pullPaperForm.target,
        status: "open",
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
        completedAt: null,
        completedBy: "",
      },
      ...current,
    ]);
    recordActivity("Created pull paper request", "Pull Paper", `${pullPaperForm.target}: ${pullPaperForm.details.trim().slice(0, 80)}`);
    setPullPaperForm(EMPTY_PULL_PAPER_FORM);
  }

  function submitSuppliesRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!suppliesForm.details.trim()) return;

    setSuppliesRequests((current) => [
      {
        id: makeId("supply"),
        details: suppliesForm.details.trim(),
        status: "open",
        attachments: suppliesDraftAttachments,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
        completedAt: null,
        completedBy: "",
      },
      ...current,
    ]);
    recordActivity("Created supplies request", "Supplies", suppliesForm.details.trim().slice(0, 100));
    setSuppliesForm(EMPTY_SUPPLIES_FORM);
    setSuppliesDraftAttachments([]);
  }

  function markPullPaperRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    const request = pullPaperRequests.find((item) => item.id === requestId);
    setPullPaperRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedBy: currentUser.username,
            }
          : request
      )
    );
    if (request) {
      recordActivity("Completed pull paper request", "Pull Paper", `${request.target}: ${safeText(request.details).slice(0, 80)}`);
    }
  }

  function deletePullPaperRequest(requestId) {
    if (!userCanEdit) return;
    const request = pullPaperRequests.find((item) => item.id === requestId);
    setPullPaperRequests((current) => current.filter((request) => request.id !== requestId));
    if (request) {
      recordActivity("Deleted pull paper request", "Pull Paper", `${request.target}: ${safeText(request.details).slice(0, 80)}`);
    }
  }

  function markSuppliesRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    const request = suppliesRequests.find((item) => item.id === requestId);
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedBy: currentUser.username,
            }
          : request
      )
    );
    if (request) {
      recordActivity("Completed supplies request", "Supplies", safeText(request.details).slice(0, 100));
    }
  }

  function deleteSuppliesRequest(requestId) {
    if (!userCanEdit) return;
    const request = suppliesRequests.find((item) => item.id === requestId);
    setSuppliesRequests((current) => current.filter((request) => request.id !== requestId));
    if (request?.attachments?.length) {
      deleteStoredAttachments(request.attachments);
    }
    if (request) {
      recordActivity("Deleted supplies request", "Supplies", safeText(request.details).slice(0, 100));
    }
  }

  async function addSuppliesRequestAttachments(requestId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: [...(request.attachments || []), ...attachments],
            }
          : request
      )
    );
    recordActivity("Added supplies request files", "Supplies", `${attachments.length} file(s) attached to a supplies request.`);
  }

  function removeSuppliesRequestAttachment(requestId, attachmentId) {
    if (!userCanEdit) return;
    const request = suppliesRequests.find((item) => item.id === requestId);
    const attachment = request?.attachments?.find((item) => item.id === attachmentId);
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: (request.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : request
      )
    );
    deleteStoredAttachments(attachment ? [attachment] : []);
  }

  function markRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    const request = requests.find((item) => item.id === requestId);
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedByAccount: currentUser.username,
            }
          : request
      )
    );
    if (request) {
      recordActivity("Completed request", "Requests", `${request.customer} ${request.jobNumber} was marked done.`);
    }
  }

  function deleteRequest(requestId) {
    if (!userCanEdit) return;
    const request = requests.find((item) => item.id === requestId);
    setRequests((current) => current.filter((request) => request.id !== requestId));
    if (request?.attachments?.length) {
      deleteStoredAttachments(request.attachments);
    }
    if (request) {
      recordActivity("Deleted request", "Requests", `${request.customer} ${request.jobNumber} was deleted.`);
    }
  }

  async function addRequestAttachments(requestId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: [...(request.attachments || []), ...attachments],
            }
          : request
      )
    );
    recordActivity("Added request files", "Requests", `${attachments.length} file(s) attached to an open request.`);
  }

  function removeRequestAttachment(requestId, attachmentId) {
    if (!userCanEdit) return;
    const request = requests.find((item) => item.id === requestId);
    const attachment = request?.attachments?.find((item) => item.id === attachmentId);
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: (request.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : request
      )
    );
    deleteStoredAttachments(attachment ? [attachment] : []);
  }

  function toggleShipmentJob(jobId) {
    setSelectedShipmentJobs((current) =>
      current.includes(jobId) ? current.filter((value) => value !== jobId) : [...current, jobId]
    );
  }

  function toggleShipQueueJob(jobId) {
    setSelectedShipQueueJobs((current) =>
      current.includes(jobId) ? current.filter((value) => value !== jobId) : [...current, jobId]
    );
  }

  function assignShipDateToFinishedJobs() {
    if (!userCanEdit) return;
    if (!selectedShipQueueJobs.length || !shipDateDraft) return;
    const targetIds = new Set(selectedShipQueueJobs);
    setAssignments((current) =>
      current.map((assignment) => {
        if (
          !targetIds.has(assignment.jobId) ||
          (assignment.kind !== "press" && assignment.kind !== "finish-only") ||
          assignment.status !== "finished"
        ) {
          return assignment;
        }
        return {
          ...assignment,
          shipDate: shipDateDraft,
        };
      })
    );
    setSelectedShipDate(shipDateDraft);
    recordActivity("Assigned ship date", "Shipping", `${selectedShipQueueJobs.length} finished jobs were moved to ship on ${shipDateDraft}.`, {
      shipDate: shipDateDraft,
      jobCount: selectedShipQueueJobs.length,
    });
    setSelectedShipQueueJobs([]);
  }

  function dismissFinishedJobFromShippingQueue(jobId) {
    if (!userCanEdit) return;
    const job = jobMap.get(jobId);
    setAssignments((current) =>
      current.map((assignment) => {
        if (
          assignment.jobId !== jobId ||
          (assignment.kind !== "press" && assignment.kind !== "finish-only") ||
          assignment.status !== "finished"
        ) {
          return assignment;
        }
        return {
          ...assignment,
          excludeFromShipping: true,
        };
      })
    );
    if (selectedShipQueueJobs.includes(jobId)) {
      setSelectedShipQueueJobs((current) => current.filter((value) => value !== jobId));
    }
    if (job) {
      recordActivity(
        "Removed finished job from ship queue",
        "Shipping",
        `${job.customerName} ${job.number} was removed from the assign ship date queue.`,
        { jobId }
      );
    }
  }

  function addShipmentMethod() {
    if (!userCanEdit) return;
    const method = safeText(newShipmentMethod);
    if (!method) return;
    if (shipmentMethods.some((item) => comparableUsername(item) === comparableUsername(method))) return;
    setShipmentMethods((current) => [...current, method]);
    setShipmentForm((current) => ({ ...current, method }));
    setNewShipmentMethod("");
    recordActivity("Added shipment method", "Shipping", `${method} is now available as a shipment method.`);
  }

  function removeShipmentMethod(methodToRemove) {
    if (!userCanEdit) return;
    if (shipmentMethods.length <= 1) return;
    setShipmentMethods((current) => {
      const next = current.filter((method) => method !== methodToRemove);
      return next.length ? next : current;
    });
    setShipmentForm((current) => {
      if (current.method !== methodToRemove) return current;
      const fallback = shipmentMethods.find((method) => method !== methodToRemove) || current.method;
        return { ...current, method: fallback };
      });
    recordActivity("Removed shipment method", "Shipping", `${methodToRemove} was removed from shipment methods.`);
  }

  function createShipmentGroup(event) {
    event.preventDefault();
    if (!userCanEdit) return;
    if (!selectedShipmentJobs.length) return;
    const jobItems = selectedShipmentJobs
      .map((jobId) => {
        const job = jobMap.get(jobId);
        if (!job) return null;
        return snapshotJobForShipment(job, finishedMetaByJobId.get(jobId));
      })
      .filter(Boolean);

    setShipmentGroups((current) => [
      {
        id: makeId("ship"),
        label: shipmentForm.label || `${shipmentForm.method} shipment`,
        method: shipmentForm.method,
        packageCount: parseNumber(shipmentForm.packageCount),
        packageType: shipmentForm.packageType,
        totalCost: parseCurrency(shipmentForm.totalCost),
        billAmount: parseCurrency(shipmentForm.billAmount),
        notes: shipmentForm.notes,
        shipDate: shipmentForm.shipDate,
        createdAt: new Date().toISOString(),
        createdBy: currentUser?.username || "",
        attachments: shipmentDraftAttachments,
        jobItems,
      },
      ...current,
    ]);

    setSelectedShipmentJobs([]);
    setShipmentForm({ ...EMPTY_SHIPMENT_FORM, shipDate: selectedShipDate });
    setShipmentDraftAttachments([]);
    recordActivity(
      "Created shipment group",
      "Shipping",
      `${shipmentForm.label || `${shipmentForm.method} shipment`} was created with ${jobItems.length} jobs for ${shipmentForm.shipDate}.`,
      {
        method: shipmentForm.method,
        shipDate: shipmentForm.shipDate,
        jobCount: jobItems.length,
      }
    );
  }

  function deleteShipmentGroup(groupId) {
    if (!userCanEdit) return;
    const group = shipmentGroups.find((item) => item.id === groupId);
    setShipmentGroups((current) => current.filter((group) => group.id !== groupId));
    if (group?.attachments?.length) {
      deleteStoredAttachments(group.attachments);
    }
    if (group) {
      recordActivity("Deleted shipment group", "Shipping", `${group.label} on ${group.shipDate} was deleted.`);
    }
  }

  async function addShipmentGroupAttachments(groupId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setShipmentGroups((current) =>
      current.map((group) =>
        group.id === groupId
          ? {
              ...group,
              attachments: [...(group.attachments || []), ...attachments],
            }
          : group
        )
    );
    recordActivity("Added shipment files", "Shipping", `${attachments.length} file(s) were added to a shipment group.`);
  }

  function removeShipmentGroupAttachment(groupId, attachmentId) {
    if (!userCanEdit) return;
    const group = shipmentGroups.find((item) => item.id === groupId);
    const attachment = group?.attachments?.find((item) => item.id === attachmentId);
    setShipmentGroups((current) =>
      current.map((group) =>
        group.id === groupId
          ? {
              ...group,
              attachments: (group.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : group
      )
    );
    deleteStoredAttachments(attachment ? [attachment] : []);
    if (attachment) {
      recordActivity("Removed shipment file", "Shipping", `${attachment.name} was removed from a shipment group.`);
    }
  }

  function updateJobRecommendedPress(jobId, press) {
    if (!userCanEdit) return;
    const nextPress = normalizePressValue(press);
    if (!PRESS_ORDER.includes(nextPress)) return;
    const job = jobMap.get(jobId);
    if (job?.press === nextPress) return;
    setJobs((current) =>
      current.map((job) =>
        job.id === jobId
          ? {
              ...job,
              press: nextPress,
            }
          : job
      )
    );
    if (job) {
      recordActivity(
        "Updated recommended press",
        "Scheduler",
        `${job.customerName} ${job.number} changed from Press ${job.press || "-"} to Press ${nextPress}.`,
        { jobId, press: nextPress }
      );
    }
  }

  function updateJobHoldState(jobId, holdActive) {
    if (!userCanEdit) return;
    const job = jobMap.get(jobId);
    setJobs((current) =>
      current.map((item) =>
        item.id === jobId
          ? {
              ...item,
              holdActive: !!holdActive,
            }
          : item
      )
    );
    if (job) {
      recordActivity(
        holdActive ? "Marked job hold" : "Cleared job hold",
        "Scheduler",
        holdActive
          ? `${job.customerName} ${job.number} was marked on hold.`
          : `${job.customerName} ${job.number} was taken off hold.`,
        { jobId, holdActive: !!holdActive }
      );
    }
  }

  function updateJobHoldNote(jobId, holdNote) {
    if (!userCanEdit) return;
    setJobs((current) =>
      current.map((item) =>
        item.id === jobId
          ? {
              ...item,
              holdNote,
            }
          : item
      )
    );
  }

  function addNote(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    const text = safeText(noteForm.text);
    if (!text) return;
    setNotes((current) => [
      {
        id: makeId("note"),
        ownerUsername: currentUser.username,
        text,
        completed: false,
        createdAt: new Date().toISOString(),
        completedAt: "",
      },
      ...current,
    ]);
    setNoteForm(EMPTY_NOTE_FORM);
  }

  function toggleNote(noteId) {
    if (!currentUser || !userCanEdit) return;
    setNotes((current) =>
      current.map((note) =>
        note.id === noteId
          ? {
              ...note,
              completed: !note.completed,
              completedAt: note.completed ? "" : new Date().toISOString(),
            }
          : note
      )
    );
  }

  function deleteNote(noteId) {
    if (!currentUser || !userCanEdit) return;
    setNotes((current) => current.filter((note) => note.id !== noteId));
  }

  function toggleUserTab(userId, tab) {
    if (!userCanManageUsers) return;
    const user = users.find((item) => item.id === userId);
    setUsers((current) =>
      current.map((user) => {
        if (user.id !== userId) return user;
        const nextTabs = user.tabs.includes(tab)
          ? user.tabs.filter((value) => value !== tab)
          : [...user.tabs, tab];
        return { ...user, tabs: normalizeUserTabs(nextTabs, user.role, user.isAdmin, user.canManageUsers) };
      })
    );
    if (user) {
      recordActivity("Changed tab visibility", "Users", `${tab} visibility was updated for ${user.username}.`, {
        userId,
        tab,
      });
    }
  }

  function updateUserAccess(userId, updates) {
    if (!userCanManageUsers) return;
    const target = users.find((user) => user.id === userId);
    if (!target) return;
    if (Object.prototype.hasOwnProperty.call(updates, "canManageUsers")) {
      const managers = users.filter((user) => user.canManageUsers);
      if (target.canManageUsers && !updates.canManageUsers && managers.length <= 1) {
        window.alert("Keep at least one user with admin access.");
        return;
      }
    }
    setUsers((current) => {
      return current.map((user) => {
        if (user.id !== userId) return user;
        const nextCanManageUsers = Object.prototype.hasOwnProperty.call(updates, "canManageUsers")
          ? !!updates.canManageUsers
          : user.canManageUsers;
        const nextRole = nextCanManageUsers ? "Management" : user.role === "Management" ? "Warehouse/Shipper" : user.role;
        return {
          ...user,
          ...updates,
          role: nextRole,
          isAdmin: nextCanManageUsers,
          canManageUsers: nextCanManageUsers,
          accessMode: Object.prototype.hasOwnProperty.call(updates, "accessMode")
            ? normalizeAccessMode(updates.accessMode, nextRole, nextCanManageUsers)
            : user.accessMode,
          tabs: normalizeUserTabs(
            Object.prototype.hasOwnProperty.call(updates, "tabs") ? updates.tabs : user.tabs,
            nextRole,
            nextCanManageUsers,
            nextCanManageUsers
          ),
        };
      });
    });
    recordActivity("Updated user access", "Users", `Access settings were updated for ${target.username}.`, {
      userId,
      updates,
    });
  }

  function openShipmentEmailDraft() {
    if (!userCanEdit) return;
    const recipients = safeText(shipmentEmailForm.recipients);
    const cc = safeText(shipmentEmailForm.cc);
    const queryParts = [
      `subject=${encodeURIComponent(shipmentEmailDraft.subject)}`,
      `body=${encodeURIComponent(shipmentEmailDraft.body)}`,
    ];
    if (cc) {
      queryParts.push(`cc=${encodeURIComponent(cc)}`);
    }
    window.location.href = `mailto:${encodeURIComponent(recipients)}?${queryParts.join("&")}`;
  }

  function logShipmentEmail() {
    if (!currentUser || !selectedShipDate || !userCanEdit) return;
    setShipmentEmailLogs((current) => [
      {
        id: makeId("email"),
        shipDate: selectedShipDate,
        recipients: safeText(shipmentEmailForm.recipients),
        cc: safeText(shipmentEmailForm.cc),
        subject: shipmentEmailDraft.subject,
        body: shipmentEmailDraft.body,
        jobCount: shipmentEmailDraft.jobCount,
        totalCost: shipmentEmailDraft.totalCost,
        totalBill: shipmentEmailDraft.totalBill,
        methods: shipmentEmailDraft.methods,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
      ...current,
    ]);
    recordActivity(
      "Logged shipment email",
      "Shipping",
      `Shipment email for ${selectedShipDate} was marked sent to ${safeText(shipmentEmailForm.recipients) || "the saved group"}${safeText(shipmentEmailForm.cc) ? ` with cc ${safeText(shipmentEmailForm.cc)}` : ""}.`,
      { shipDate: selectedShipDate, recipients: safeText(shipmentEmailForm.recipients), cc: safeText(shipmentEmailForm.cc) }
    );
  }

  function deleteShipmentEmailLog(logId) {
    if (!userCanEdit) return;
    const log = shipmentEmailLogs.find((item) => item.id === logId);
    setShipmentEmailLogs((current) => current.filter((entry) => entry.id !== logId));
    if (log) {
      recordActivity(
        "Deleted shipment email log",
        "Shipping",
        `Shipment email history for ${log.shipDate} was deleted.`,
        { logId, shipDate: log.shipDate }
      );
    }
  }

  function applyShipmentEmailGroup(groupId) {
    const group = shipmentEmailGroups.find((item) => item.id === groupId);
    setShipmentEmailForm({
      groupId,
      recipients: group?.recipients || "",
      cc: group?.cc || "",
    });
  }

  function saveShipmentEmailGroup() {
    if (!userCanEdit) return;
    const name = safeText(shipmentEmailGroupForm.name);
    const recipients = safeText(shipmentEmailGroupForm.recipients || shipmentEmailForm.recipients);
    const cc = safeText(shipmentEmailGroupForm.cc || shipmentEmailForm.cc);
    if (!name || (!recipients && !cc)) return;

    setShipmentEmailGroups((current) => {
      const existing = current.find((group) => comparableUsername(group.name) === comparableUsername(name));
      if (existing) {
        return current.map((group) => (group.id === existing.id ? { ...group, name, recipients, cc } : group));
      }
      return [...current, { id: makeId("email-group"), name, recipients, cc }];
    });

    setShipmentEmailForm((current) => ({ ...current, recipients, cc }));
    setShipmentEmailGroupForm(EMPTY_EMAIL_GROUP_FORM);
    recordActivity("Saved email group", "Shipping", `${name} was saved as a shipment email group.`);
  }

  function removeShipmentEmailGroup(groupId) {
    if (!userCanEdit) return;
    const group = shipmentEmailGroups.find((item) => item.id === groupId);
    setShipmentEmailGroups((current) => current.filter((group) => group.id !== groupId));
    setShipmentEmailForm((current) =>
      current.groupId === groupId ? { groupId: "", recipients: "", cc: "" } : current
    );
    if (group) {
      recordActivity("Removed email group", "Shipping", `${group.name} was removed from shipment email groups.`);
    }
  }

  function switchAuthView(nextView) {
    setAuthView(nextView);
    setLoginError("");
    setRegisterError("");
    setRegisterSuccess("");
  }

  function enterDemoWorkspace() {
    updateDemoWorkspaceUrl(true);
    const demoSnapshot = buildDemoSharedSnapshot();
    const normalized = applySharedStateSnapshot(demoSnapshot);
    lastSharedSnapshotRef.current = JSON.stringify(buildSharedSnapshot(normalized));
    setWorkspaceMode("demo");
    setSyncStatus("Demo mode");
    setLastSyncAt("");
    setCurrentUsername(DEMO_DEFAULT_USERNAME);
    setSessionExpiresAt(new Date(Date.now() + LOGIN_SESSION_DURATION_MS).toISOString());
    setLoginForm({ username: DEMO_DEFAULT_USERNAME, password: DEMO_DEFAULT_PASSWORD });
    setLoginError("");
    setRegisterError("");
    setRegisterSuccess("");
    setAuthView("login");
    setActiveTab("Scheduler");
  }

  function exitDemoWorkspace() {
    if (workspaceMode !== "demo") return;
    updateDemoWorkspaceUrl(false);
    window.location.reload();
  }

  function handleLogin(event) {
    event.preventDefault();
    const match = users.find(
      (user) =>
        comparableUsername(user.username) === comparableUsername(loginForm.username) &&
        user.password === loginForm.password
    );
    if (!match) {
      setLoginError("Invalid username or password.");
      return;
    }
    setCurrentUsername(match.username);
    setSessionExpiresAt(new Date(Date.now() + LOGIN_SESSION_DURATION_MS).toISOString());
    setLoginForm(EMPTY_LOGIN_FORM);
    setLoginError("");
    setRegisterSuccess("");
    setAuthView("login");
    setActiveTab("Scheduler");
  }

  function submitRegistrationRequest(event) {
    event.preventDefault();
    const username = safeText(registerForm.username);
    const password = safeText(registerForm.password);

    if (!username || !password) {
      setRegisterError("Enter both a name and password.");
      setRegisterSuccess("");
      return;
    }

    const hasUser = users.some((user) => comparableUsername(user.username) === comparableUsername(username));
    if (hasUser) {
      setRegisterError("That username already exists.");
      setRegisterSuccess("");
      return;
    }

    const hasPendingRequest = registrationRequests.some(
      (request) =>
        comparableUsername(request.username) === comparableUsername(username) && request.status === "pending"
    );
    if (hasPendingRequest) {
      setRegisterError("That account request is already waiting for approval.");
      setRegisterSuccess("");
      return;
    }

    setRegistrationRequests((current) => [
      ...current,
      {
        id: makeId("registration"),
        username,
        password,
        status: "pending",
        createdAt: new Date().toISOString(),
        createdBy: username,
        approvedAt: "",
        approvedBy: "",
        deniedAt: "",
        deniedBy: "",
      },
    ]);
    setRegisterForm(EMPTY_REGISTER_FORM);
    setRegisterError("");
    setRegisterSuccess("Registration request sent. A manager will need to approve it.");
    setAuthView("login");
    setActivityLog((current) =>
      appendActivityEntry(current, {
        actor: username,
        action: "Requested account access",
        scope: "Users",
        description: `${username} submitted a registration request.`,
      })
    );
  }

  function handleLogout() {
    setCurrentUsername("");
    setSessionExpiresAt("");
    setAuthView("login");
    setActiveTab("Scheduler");
  }

  function createUser(event) {
    event.preventDefault();
    if (!userCanManageUsers) return;
    const username = safeText(userForm.username);
    const password = safeText(userForm.password);
    if (!username || !password) return;
    const exists = users.some((user) => comparableUsername(user.username) === comparableUsername(username));
    if (exists) {
      window.alert("That username already exists.");
      return;
    }
    setUsers((current) => [
      ...current,
      {
        id: makeId("user"),
        username,
        password,
        role: userForm.canManageUsers ? "Management" : "Warehouse/Shipper",
        accessMode: normalizeAccessMode(userForm.accessMode, userForm.canManageUsers ? "Management" : "Warehouse/Shipper", userForm.canManageUsers),
        tabs: normalizeUserTabs(userForm.tabs, userForm.canManageUsers ? "Management" : "Warehouse/Shipper", userForm.canManageUsers, userForm.canManageUsers),
        canManageUsers: !!userForm.canManageUsers,
        isAdmin: !!userForm.canManageUsers,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
    ]);
    setUserForm({ ...EMPTY_USER_FORM, tabs: [...EMPTY_USER_FORM.tabs] });
    recordActivity("Created user", "Users", `${username} was added as a new user.`);
  }

  function updateUserPassword(userId) {
    if (!userCanManageUsers) return;
    const nextPassword = safeText(userPasswordDrafts[userId]);
    if (!nextPassword) return;
    const user = users.find((item) => item.id === userId);
    setUsers((current) =>
      current.map((user) => (user.id === userId ? { ...user, password: nextPassword } : user))
    );
    setUserPasswordDrafts((current) => ({ ...current, [userId]: "" }));
    if (user) {
      recordActivity("Reset user password", "Users", `Password was updated for ${user.username}.`);
    }
  }

  function renameUser(userId) {
    if (!userCanManageUsers) return;
    const target = users.find((user) => user.id === userId);
    const nextUsername = safeText(userUsernameDrafts[userId]);
    if (!target || !nextUsername || comparableUsername(target.username) === comparableUsername(nextUsername)) return;
    const exists = users.some(
      (user) => user.id !== userId && comparableUsername(user.username) === comparableUsername(nextUsername)
    );
    if (exists) {
      window.alert("That username already exists.");
      return;
    }

    const previousUsername = target.username;
    const isSameUser = (value) => comparableUsername(value) === comparableUsername(previousUsername);

    setUsers((current) => current.map((user) => (user.id === userId ? { ...user, username: nextUsername } : user)));
    setRequests((current) =>
      current.map((request) => ({
        ...request,
        assignedToAccount: isSameUser(request.assignedToAccount) ? nextUsername : request.assignedToAccount,
        createdByAccount: isSameUser(request.createdByAccount) ? nextUsername : request.createdByAccount,
        completedByAccount: isSameUser(request.completedByAccount) ? nextUsername : request.completedByAccount,
      }))
    );
    setPullPaperRequests((current) =>
      current.map((request) => ({
        ...request,
        createdBy: isSameUser(request.createdBy) ? nextUsername : request.createdBy,
        completedBy: isSameUser(request.completedBy) ? nextUsername : request.completedBy,
      }))
    );
    setSuppliesRequests((current) =>
      current.map((request) => ({
        ...request,
        createdBy: isSameUser(request.createdBy) ? nextUsername : request.createdBy,
        completedBy: isSameUser(request.completedBy) ? nextUsername : request.completedBy,
      }))
    );
    setNotes((current) =>
      current.map((note) => ({
        ...note,
        ownerUsername: isSameUser(note.ownerUsername) ? nextUsername : note.ownerUsername,
      }))
    );
    setAssignments((current) =>
      current.map((assignment) => ({
        ...assignment,
        finishedBy: isSameUser(assignment.finishedBy) ? nextUsername : assignment.finishedBy,
      }))
    );
    setShipmentGroups((current) =>
      current.map((group) => ({
        ...group,
        createdBy: isSameUser(group.createdBy) ? nextUsername : group.createdBy,
        jobItems: Array.isArray(group.jobItems)
          ? group.jobItems.map((item) => ({
              ...item,
              finishedBy: isSameUser(item.finishedBy) ? nextUsername : item.finishedBy,
            }))
          : group.jobItems,
      }))
    );
    setShipmentEmailLogs((current) =>
      current.map((log) => ({
        ...log,
        createdBy: isSameUser(log.createdBy) ? nextUsername : log.createdBy,
      }))
    );
    setRegistrationRequests((current) =>
      current.map((request) => ({
        ...request,
        approvedBy: isSameUser(request.approvedBy) ? nextUsername : request.approvedBy,
        deniedBy: isSameUser(request.deniedBy) ? nextUsername : request.deniedBy,
      }))
    );
    setActivityLog((current) =>
      current.map((entry) => ({
        ...entry,
        actor: isSameUser(entry.actor) ? nextUsername : entry.actor,
      }))
    );
    if (isSameUser(currentUsername)) {
      setCurrentUsername(nextUsername);
    }
    setUserUsernameDrafts((current) => ({ ...current, [userId]: "" }));
    recordActivity("Renamed user", "Users", `${previousUsername} was renamed to ${nextUsername}.`, {
      userId,
      previousUsername,
      nextUsername,
    });
  }

  function deleteUser(userId) {
    if (!userCanManageUsers) return;
    const target = users.find((user) => user.id === userId);
    if (!target) return;
    if (comparableUsername(target.username) === comparableUsername(currentUsername)) {
      window.alert("You cannot delete the account you are currently signed in with.");
      return;
    }
    if (target.canManageUsers) {
      const managers = users.filter((user) => user.canManageUsers);
      if (managers.length <= 1) {
        window.alert("Keep at least one user with admin access.");
        return;
      }
    }
    const confirmed = window.confirm(`Delete the ${target.username} account? This will remove personal notes and clear open request assignments for that user.`);
    if (!confirmed) return;

    const isSameUser = (value) => comparableUsername(value) === comparableUsername(target.username);
    setUsers((current) => current.filter((user) => user.id !== userId));
    setRequests((current) =>
      current.map((request) => ({
        ...request,
        assignedToAccount: isSameUser(request.assignedToAccount) ? "" : request.assignedToAccount,
      }))
    );
    setNotes((current) => current.filter((note) => !isSameUser(note.ownerUsername)));
    setUserPasswordDrafts((current) => {
      const next = { ...current };
      delete next[userId];
      return next;
    });
    setUserUsernameDrafts((current) => {
      const next = { ...current };
      delete next[userId];
      return next;
    });
    recordActivity("Deleted user", "Users", `${target.username} was deleted from the scheduler.`, {
      userId,
      username: target.username,
    });
  }

  function approveRegistrationRequest(requestId) {
    if (!userCanManageUsers) return;
    const request = registrationRequests.find((item) => item.id === requestId);
    if (!request) return;

    const exists = users.some(
      (user) => comparableUsername(user.username) === comparableUsername(request.username)
    );
    if (exists) {
      window.alert("That username already exists. Delete or rename the pending request first.");
      return;
    }

    setUsers((current) => [
      ...current,
      {
        id: makeId("user"),
        username: request.username,
        password: request.password,
        role: "Warehouse/Shipper",
        accessMode: "edit",
        tabs: normalizeUserTabs(EMPTY_USER_FORM.tabs, "Warehouse/Shipper", false, false),
        canManageUsers: false,
        isAdmin: false,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
    ]);
    setRegistrationRequests((current) => current.filter((item) => item.id !== requestId));
    recordActivity("Approved registration", "Users", `${request.username} was approved and added as a user.`);
  }

  function denyRegistrationRequest(requestId) {
    if (!userCanManageUsers) return;
    const request = registrationRequests.find((item) => item.id === requestId);
    setRegistrationRequests((current) => current.filter((item) => item.id !== requestId));
    if (request) {
      recordActivity("Denied registration", "Users", `${request.username} was denied account access.`);
    }
  }

  if (!isReady) {
    return (
      <div className="min-h-screen bg-stone-100 p-6 text-stone-900">
        <div className="mx-auto max-w-xl rounded-3xl border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-8 text-center shadow-sm shadow-stone-300/40">
          <div className="text-lg font-semibold">Loading scheduler...</div>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <LoginScreen
        authView={authView}
        loginForm={loginForm}
        loginError={loginError}
        registerForm={registerForm}
        registerError={registerError}
        registerSuccess={registerSuccess}
        users={users}
        onChangeLogin={setLoginForm}
        onChangeRegister={setRegisterForm}
        onChangeView={switchAuthView}
        onSubmitLogin={handleLogin}
        onSubmitRegister={submitRegistrationRequest}
        workspaceMode={workspaceMode}
        onEnterDemo={enterDemoWorkspace}
        onExitDemo={exitDemoWorkspace}
      />
    );
  }

  const schedulerCards = [
    ["TXT open", summary.txtOpen],
    ["TXT closed", summary.txtClosed],
    ["Marked finished", summary.markedFinished],
    ["Open requests", summary.openRequests],
    ["Scheduled jobs", summary.scheduledJobs],
    ["Ship groups", summary.shipGroupsOnDate],
  ];

  const tabBadges = {
    Notes: userNotes.filter((note) => !note.completed).length,
    "Open Requests": openRequests.length,
    "Pull Paper Request": openPullPaperRequests.length,
    "Supplies Request": openSuppliesRequests.length,
  };

  return (
    <div className="min-h-screen bg-stone-100 text-stone-900">
      <div className="mx-auto max-w-[1900px] p-4 md:p-6">
        <div className="mb-6 rounded-[2rem] border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-5 shadow-sm shadow-stone-300/40">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-emerald-900">
                {workspaceMode === "demo" ? "Demo board" : "Production board"}
              </p>
              <h1 className="mt-1 text-2xl font-semibold tracking-tight">Label Traxx Scheduler</h1>
              <p className="mt-2 max-w-3xl text-sm text-stone-700">
                {workspaceMode === "demo"
                  ? `Logged in as ${currentUser.username} in the demo workspace. Everything here is sample data only for presentation use.`
                  : `Logged in as ${currentUser.username}. Request history, attachments, and completed work are now tied to user accounts.`}
              </p>
            </div>
            <div className="flex flex-col gap-3 xl:items-end">
              <div className="w-full overflow-x-auto pb-1 xl:max-w-[72vw]">
                <div className="flex min-w-max gap-2">
                  {tabs.map((tab) => (
                    <button
                      key={tab}
                      onClick={() => setActiveTab(tab)}
                      className={`rounded-2xl px-4 py-2 text-sm font-medium transition ${
                        activeTab === tab
                          ? "bg-emerald-900 text-stone-50 shadow-sm"
                          : "border border-stone-300 bg-stone-50 text-stone-700 hover:bg-stone-100"
                      }`}
                    >
                      <span>{tab}</span>
                      {tabBadges[tab] > 0 && (
                        <span
                          className={`ml-2 rounded-full px-2 py-0.5 text-[11px] font-semibold ${
                            activeTab === tab ? "bg-white/15 text-stone-50" : "bg-stone-200 text-stone-800"
                          }`}
                        >
                          {tabBadges[tab]}
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              </div>
              <div className="flex items-center gap-2 text-sm">
                <span className={`rounded-full px-3 py-2 ${syncTone(syncStatus)}`}>
                  {syncStatus}
                  {lastSyncAt ? ` • ${formatDateTime(lastSyncAt)}` : ""}
                </span>
                <span className="rounded-full bg-stone-200 px-3 py-2 text-stone-800">
                  {currentUserRole} / {currentUserAccessMode === "edit" ? "Edit" : "View only"}: {currentUser.username}
                </span>
                {workspaceMode === "demo" && (
                  <button
                    type="button"
                    onClick={exitDemoWorkspace}
                    className="rounded-2xl border border-sky-200 bg-sky-50 px-3 py-2 text-sky-900"
                  >
                    Exit demo
                  </button>
                )}
                <button onClick={handleLogout} className="rounded-2xl border border-stone-300 bg-stone-50 px-3 py-2 text-stone-800">
                  Log out
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="mb-6 grid gap-3 md:grid-cols-3 xl:grid-cols-6">
          {schedulerCards.map(([label, value]) => (
            <div key={label} className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
              <div className="text-xs uppercase tracking-[0.16em] text-stone-600">{label}</div>
              <div className="mt-2 text-2xl font-semibold tracking-tight">{value}</div>
            </div>
          ))}
        </div>

        {activeTab === "Today" && (
          <div className="space-y-4">
            <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-6">
              {todayDashboardStats.map((item) => (
                <div key={item.label} className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                  <div className="text-xs uppercase tracking-[0.16em] text-stone-600">{item.label}</div>
                  <div className="mt-2 text-2xl font-semibold tracking-tight">{item.value}</div>
                </div>
              ))}
            </div>

            <div className="grid gap-4 2xl:grid-cols-[minmax(0,1.2fr)_minmax(0,0.8fr)]">
              <div className="space-y-4">
                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <div className="text-sm font-semibold">Today&apos;s board</div>
                      <div className="text-xs text-stone-600">
                        Everything scheduled for {todayDateKey}, including manual blocks and multi-day work.
                      </div>
                    </div>
                    <button
                      onClick={() => {
                        setWeekStart(startOfWeek(new Date(`${todayDateKey}T12:00:00`)));
                        setActiveTab("Scheduler");
                      }}
                      className="rounded-2xl border border-stone-300 bg-white px-4 py-2 text-sm text-stone-800"
                    >
                      Open scheduler
                    </button>
                  </div>
                  <div className="space-y-3">
                    {todayScheduledItems.map((item) => (
                      <div key={item.assignment.id} className="rounded-2xl border border-stone-300 bg-white p-3">
                        <div className="mb-2 flex flex-wrap items-center gap-2">
                          <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">
                            {formatPressLabel(item.press)}
                          </span>
                          <span
                            className={`rounded-full px-2 py-1 text-[11px] font-medium ${
                              item.assignment.status === "finished" ? statusTone("done") : statusTone(item.assignment.status)
                            }`}
                          >
                            {item.assignment.kind === "manual"
                              ? "manual block"
                              : item.assignment.status === "finished"
                                ? "done"
                                : item.assignment.status}
                          </span>
                        </div>
                        <CompactScheduleCard
                          job={item.job}
                          assignment={item.assignment}
                          finishMeta={item.job ? finishedMetaByJobId.get(item.job.id) : null}
                          density="detailed"
                          selected={item.job ? selectedJobId === item.job.id : false}
                          onSelect={item.job ? () => selectJob(item.job.id) : undefined}
                          draggable={false}
                        />
                      </div>
                    ))}
                    {!todayScheduledItems.length && (
                      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                        Nothing is on today&apos;s board yet.
                      </div>
                    )}
                  </div>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <div className="text-sm font-semibold">Jobs due today</div>
                      <div className="text-xs text-stone-600">Open work due today that has not been marked finished yet.</div>
                    </div>
                    <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{jobsDueToday.length}</div>
                  </div>
                  <div className="space-y-3">
                    {jobsDueToday.map((job) => (
                      <div key={job.id} className="rounded-2xl border border-stone-300 bg-white p-3">
                        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                          <div className="min-w-0">
                            <div className="text-sm font-semibold text-stone-900">
                              {job.customerName} {job.number}
                            </div>
                            <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                            <div className="mt-2 text-[11px] text-stone-600">
                              {formatPressLabel(job.press || "-")} · EST {job.estPressTime.toFixed(2)} hrs
                            </div>
                          </div>
                          <button
                            onClick={() => {
                              setActiveTab("Scheduler");
                              selectJob(job.id, true);
                            }}
                            className="rounded-2xl border border-stone-300 bg-stone-50 px-3 py-2 text-sm text-stone-800"
                          >
                            Open
                          </button>
                        </div>
                      </div>
                    ))}
                    {!jobsDueToday.length && (
                      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                        No open jobs are due today.
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="space-y-4">
                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Queue snapshot</div>
                    <div className="text-xs text-stone-600">Quick links into the work that still needs attention.</div>
                  </div>
                  <div className="grid gap-3">
                    {[
                      canAccessTab(currentUser, "Open Requests")
                        ? { tab: "Open Requests", label: "Open requests", value: openRequests.length }
                        : null,
                      canAccessTab(currentUser, "Pull Paper Request")
                        ? { tab: "Pull Paper Request", label: "Pull paper", value: openPullPaperRequests.length }
                        : null,
                      canAccessTab(currentUser, "Supplies Request")
                        ? { tab: "Supplies Request", label: "Supplies", value: openSuppliesRequests.length }
                        : null,
                    ]
                      .filter(Boolean)
                      .map((item) => (
                        <button
                          key={item.tab}
                          onClick={() => setActiveTab(item.tab)}
                          className="flex items-center justify-between rounded-2xl border border-stone-300 bg-white px-4 py-3 text-left"
                        >
                          <span className="text-sm font-medium text-stone-800">{item.label}</span>
                          <span className="rounded-full bg-stone-200 px-2 py-1 text-xs text-stone-800">{item.value}</span>
                        </button>
                      ))}
                    <button
                      onClick={() => setActiveTab("Scheduler")}
                      className="flex items-center justify-between rounded-2xl border border-stone-300 bg-white px-4 py-3 text-left"
                    >
                      <span className="text-sm font-medium text-stone-800">Press queue</span>
                      <span className="rounded-full bg-stone-200 px-2 py-1 text-xs text-stone-800">{unscheduledJobs.length}</span>
                    </button>
                  </div>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Shipping today</div>
                    <div className="text-xs text-stone-600">See whether today&apos;s finished work has been grouped and emailed.</div>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-2">
                    <StatRow label="Finished today" value={finishedTodayJobs.length} />
                    <StatRow label="Ship groups" value={shipGroupsToday.length} />
                  </div>
                  <div className="mt-3 rounded-2xl bg-white p-4 text-sm text-stone-800">
                    <div className="font-medium">
                      {shipmentEmailSentToday ? "Daily shipment email already logged" : "Daily shipment email still needs to be logged"}
                    </div>
                    <div className="mt-1 text-xs text-stone-600">
                      {shipmentEmailSentToday
                        ? `The ${todayDateKey} shipment summary has already been marked sent.`
                        : `Nothing has been marked emailed for ${todayDateKey} yet.`}
                    </div>
                  </div>
                  {canAccessTab(currentUser, "Daily Shipment") && (
                    <button
                      onClick={() => {
                        setSelectedShipDate(todayDateKey);
                        setActiveTab("Daily Shipment");
                      }}
                      className="mt-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                    >
                      Open daily shipment
                    </button>
                  )}
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4 flex items-center justify-between">
                    <div>
                      <div className="text-sm font-semibold">My checklist</div>
                      <div className="text-xs text-stone-600">Private notes tied only to {currentUser.username}.</div>
                    </div>
                    <span className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">
                      {userNotes.filter((note) => !note.completed).length} open
                    </span>
                  </div>
                  <div className="space-y-2">
                    {userNotes.slice(0, 4).map((note) => (
                      <div key={note.id} className="rounded-2xl border border-stone-300 bg-white px-4 py-3">
                        <div className={`text-sm ${note.completed ? "text-stone-500 line-through" : "text-stone-800"}`}>
                          {note.text}
                        </div>
                      </div>
                    ))}
                    {!userNotes.length && (
                      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                        No checklist items yet.
                      </div>
                    )}
                  </div>
                  {canAccessTab(currentUser, "Notes") && (
                    <button
                      onClick={() => setActiveTab("Notes")}
                      className="mt-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                    >
                      Open notes
                    </button>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "Scheduler" && (
          <div className="space-y-4">
            {!boardFocusMode && (
              <div className="grid gap-4 2xl:grid-cols-[320px_minmax(0,1fr)]">
                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Import and controls</div>
                    <div className="text-xs text-stone-600">Imported done tickets stay informational only until a logged-in user marks them finished.</div>
                  </div>

                  <div className="grid gap-3">
                    <label className="rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                      <div className="font-medium text-stone-900">Upload Label Traxx TXT</div>
                      <input type="file" accept=".txt,.tsv,text/plain" onChange={handleUpload} className="mt-3 block w-full text-xs" />
                    </label>

                    {userCanManageUsers && (
                      <label className="rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                        <div className="flex items-center justify-between gap-3">
                          <div>
                            <div className="font-medium text-stone-900">Upload weekly schedule Excel</div>
                            <div className="mt-1 text-xs text-stone-600">Management only. Reads weekly tabs like `5-4 5-8` and ignores raw upload tabs like `5-18`.</div>
                          </div>
                          <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] text-stone-700">Admin</span>
                        </div>
                        <input type="file" accept=".xls,.xlsx,.xlsm" onChange={handleScheduleWorkbookUpload} className="mt-3 block w-full text-xs" />
                      </label>
                    )}

                    <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                      Finishing jobs and completing requests will be recorded under <span className="font-semibold">{currentUser.username}</span>.
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <button onClick={exportSchedule} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                        Export CSV
                      </button>
                      <button onClick={clearBoard} className="rounded-2xl border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-900">
                        Clear
                      </button>
                      {userCanManageUsers && (
                        <button onClick={exportScheduleWorkbook} className="col-span-2 rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                          Export Excel
                        </button>
                      )}
                    </div>

                    {canAccessTab(currentUser, "New Request") && (
                      <button
                        onClick={() => setActiveTab("New Request")}
                        className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white"
                      >
                        Create a request
                      </button>
                    )}

                    <div className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="mb-3">
                        <div className="text-sm font-medium text-stone-900">Manual schedule block</div>
                        <div className="text-xs text-stone-600">Reserve a press/day for something like Maintenance, Changeover, or Wash-up.</div>
                      </div>
                      <form onSubmit={addManualScheduleEntry} className="grid gap-3">
                        <Field
                          label="Title"
                          value={manualScheduleForm.title}
                          onChange={(value) => setManualScheduleForm((current) => ({ ...current, title: value }))}
                          placeholder="Maintenance"
                        />
                        <div>
                          <div className="mb-2 text-sm font-medium text-stone-800">Press</div>
                          <select
                            value={manualScheduleForm.press}
                            onChange={(event) => setManualScheduleForm((current) => ({ ...current, press: event.target.value }))}
                            className="w-full rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                          >
                            {PRESS_ORDER.map((press) => (
                              <option key={press} value={press}>
                                  {formatPressLabel(press)}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div>
                          <div className="mb-2 text-sm font-medium text-stone-800">Day</div>
                          <select
                            value={manualScheduleForm.dayKey}
                            onChange={(event) => setManualScheduleForm((current) => ({ ...current, dayKey: event.target.value }))}
                            className="w-full rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                          >
                            {weekColumns.map((day) => (
                              <option key={day.key} value={day.key}>
                                {day.label} - {formatShortDate(day.date)}
                              </option>
                            ))}
                          </select>
                        </div>
                        <button
                          type="submit"
                          disabled={!userCanMoveJobs}
                          className="rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800 disabled:cursor-not-allowed disabled:opacity-50"
                        >
                          Add block
                        </button>
                      </form>
                    </div>
                  </div>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                  <div className="mb-3 flex items-center justify-between">
                    <div>
                      <div className="text-sm font-semibold">Press queue</div>
                      <div className="text-xs text-stone-600">Search by ticket, then filter the queue by status, press, schedule state, or priority. Jobs stay here even after they are scheduled so you can place them on multiple days.</div>
                    </div>
                    <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{unscheduledJobs.length}</div>
                  </div>

                  {pickedUpItem && (
                    <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border border-sky-200 bg-sky-50 px-3 py-2 text-sm text-sky-900">
                      <div>
                        Picked up: <span className="font-semibold">{pickedUpItem.label || "Schedule item"}</span>. Scroll to any lane and click the lane body to place it.
                      </div>
                      <button
                        type="button"
                        onClick={() => setPickedUpItem(null)}
                        className="border border-sky-300 bg-white px-2 py-1 text-xs text-sky-900"
                      >
                        Cancel
                      </button>
                    </div>
                  )}

                  <input
                    type="text"
                    value={unscheduledSearch}
                    onChange={(event) => setUnscheduledSearch(event.target.value)}
                    placeholder="Search ticket, customer, or description"
                    className="mb-3 w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />

                  <div className="mb-3 grid gap-3">
                    <div className="flex flex-wrap gap-2">
                      {["All", "Scheduled", "Not scheduled"].map((option) => (
                        <button
                          key={option}
                          type="button"
                          onClick={() => setQueueScheduleFilter(option)}
                          className={`border px-3 py-2 text-sm ${queueScheduleFilter === option ? "border-emerald-900 bg-emerald-900 text-white" : "border-stone-300 bg-white text-stone-800"}`}
                        >
                          {option}
                        </button>
                      ))}
                    </div>
                    <div className="grid gap-2 md:grid-cols-2">
                      <select
                        value={queueStatusFilter}
                        onChange={(event) => setQueueStatusFilter(event.target.value)}
                        className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                      >
                        <option>All</option>
                        <option>Open</option>
                        <option>Done</option>
                      </select>
                      <select
                        value={queuePressFilter}
                        onChange={(event) => setQueuePressFilter(event.target.value)}
                        className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                      >
                        {queuePressOptions.map((press) => (
                          <option key={press}>{press}</option>
                        ))}
                      </select>
                    </div>
                    <div className="rounded-2xl border border-stone-300 bg-white p-3">
                      <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Priority</div>
                      <div className="flex flex-wrap gap-2">
                        {queuePriorityOptions.map((priority) => {
                          const active = queuePriorityFilters.includes(priority);
                          return (
                            <label key={priority} className={`flex items-center gap-2 border px-2 py-1 text-xs ${active ? "border-emerald-900 bg-emerald-50 text-emerald-900" : "border-stone-300 bg-stone-50 text-stone-800"}`}>
                              <input
                                type="checkbox"
                                checked={active}
                                onChange={() => toggleQueuePriority(priority)}
                                className="h-3.5 w-3.5"
                              />
                              <span>{priority}</span>
                            </label>
                          );
                        })}
                        {!queuePriorityOptions.length && <div className="text-xs text-stone-500">No priorities loaded yet.</div>}
                      </div>
                    </div>
                  </div>

                  <div className="max-h-[65vh] overflow-auto border border-stone-300 bg-white 2xl:max-h-[720px]">
                    {unscheduledJobs.length ? (
                      <table className="min-w-[1860px] border-collapse text-left text-[12px] text-stone-800">
                        <thead className="sticky top-0 z-10 bg-stone-100">
                          <tr className="border-b border-stone-300">
                            {["Number", "Customer", "Priority", "PO No.", "Description", "Press", "Ship", "Due on site", "Quantity", "Status", "Stock", "Press Time", "Main", "Scheduled On", "Mon-Fri", "Actions"].map((label) => (
                              <th key={label} className="whitespace-nowrap border-r border-stone-200 px-3 py-2 font-semibold text-stone-700 last:border-r-0">
                                {label}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {unscheduledJobs.map((job) => (
                            <PressQueueRow
                              key={job.id}
                              job={job}
                              state={deriveVisibleJobState(job.id, activePressJobIds, userFinishedJobIds)}
                              selected={selectedJobId === job.id}
                              onSelect={() => selectJob(job.id)}
                              onOpenDetails={() => selectJob(job.id, true)}
                              onFinish={userCanEdit ? () => finishJob(job.id) : undefined}
                              canMove={userCanMoveJobs}
                              scheduledAssignments={allScheduleLocationsByJobId.get(job.id) || []}
                              pressOptions={PRESS_ORDER}
                              onUpdatePress={userCanMoveJobs ? (press) => updateJobRecommendedPress(job.id, press) : undefined}
                              weekColumns={weekColumns}
                              onQuickAssign={(dayKey) => addAssignment(job.id, dayKey, PRESS_ORDER.includes(job.press) ? job.press : "Rewind")}
                              onPickUp={userCanMoveJobs ? () => pickUpQueueJob(job.id) : undefined}
                            />
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <div className="p-4 text-sm text-stone-600">
                        No open queue jobs match your search.
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            <div className="rounded-2xl border border-stone-300 bg-white p-2 shadow-sm shadow-stone-300/20">
              <div className="mb-3 flex flex-col gap-3 border border-stone-300 bg-stone-50 px-4 py-3">
                <div>
                  <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-stone-600">Week view</div>
                  <div className="mt-1 text-base font-semibold">
                    {formatShortDate(weekColumns[0]?.date)} - {formatShortDate(weekColumns[4]?.date)}
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button onClick={() => setWeekStart(addDays(weekStart, -7))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    Previous
                  </button>
                  <button onClick={() => setWeekStart(startOfWeek(new Date()))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    This week
                  </button>
                  <button onClick={() => setWeekStart(addDays(weekStart, 7))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    Next
                  </button>
                  <select
                    value={scheduleCardDensity}
                    onChange={(event) => setScheduleCardDensity(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    {SCHEDULE_DENSITY_OPTIONS.map((option) => (
                      <option key={option} value={option}>
                        {SCHEDULE_DENSITY_LABELS[option]}
                      </option>
                    ))}
                  </select>
                  <label className="flex items-center gap-2 rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    <input
                      type="checkbox"
                      checked={hideEmptyPresses}
                      onChange={(event) => setHideEmptyPresses(event.target.checked)}
                      className="h-4 w-4"
                    />
                    <span>Hide empty presses</span>
                  </label>
                  <div className="flex flex-wrap gap-2 rounded-2xl border border-stone-300 bg-white px-3 py-2">
                    {PRESS_ORDER.map((press) => {
                      const checked = visiblePressFilters.includes(press);
                      return (
                        <label key={press} className={`flex items-center gap-2 text-xs ${checked ? "text-emerald-900" : "text-stone-700"}`}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleVisiblePressFilter(press)}
                            className="h-3.5 w-3.5"
                          />
                          <span>{formatPressLabel(press)}</span>
                        </label>
                      );
                    })}
                  </div>
                  <button
                    type="button"
                    onClick={() => setBoardFocusMode((current) => !current)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    {boardFocusMode ? "Show side panels" : "Focus board"}
                  </button>
                  <input
                    type="text"
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="Search all jobs"
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
              </div>
              <div className="overflow-x-auto pb-2">
                <div
                  className="grid min-w-[1240px] gap-0 border-l border-t border-stone-300"
                  style={{ gridTemplateColumns: "160px repeat(5, minmax(0, 1fr))" }}
                >
                  <div className="border-b border-r border-stone-300 bg-stone-100 px-4 py-3">
                    <div className="text-xs uppercase tracking-[0.16em] text-stone-600">Press</div>
                    <div className="mt-1 text-sm font-semibold text-stone-900">Week lanes</div>
                  </div>
                  {weekColumns.map((day) => (
                    <div key={day.key} className="border-b border-r border-stone-300 bg-white">
                      <div className="px-3 py-1.5 text-center text-xs uppercase tracking-[0.16em] text-stone-700">{day.label}</div>
                      <div className="border-t border-stone-300 bg-sky-200 px-3 py-1 text-center text-base font-semibold text-stone-900">
                        {formatDate(day.date)}
                      </div>
                    </div>
                  ))}

                  {visibleScheduleRows.map((press) => (
                    <React.Fragment key={press}>
                      <div className="border-b border-r border-stone-300 bg-white px-4 py-3">
                        <div className="text-center text-xl font-semibold text-stone-900">{formatPressLabel(press).replace("Press ", "")}</div>
                        <div className="mt-1 text-center text-[11px] text-stone-600">Aligned across the week</div>
                      </div>
                      {weekColumns.map((day) => {
                        const laneJobs = board[day.key]?.[press] || [];
                        const totalHours = laneJobs.reduce((sum, item) => sum + (item.job?.estPressTime || 0), 0);
                        const operatorName = getPressOperator(day.key, press);
                        const dutyName = getPressDuty(day.key, press);
                        return (
                          <div
                            key={`${press}-${day.key}`}
                            onDragOver={(event) => event.preventDefault()}
                            onDrop={(event) => handleScheduleDrop(event, day.key, press)}
                            onClick={() => placePickedUpItem(day.key, press)}
                            className={`border-b border-r border-stone-300 bg-white p-0 ${pickedUpItem ? "cursor-copy" : ""} ${pickedUpItem ? "hover:bg-sky-50" : ""}`}
                          >
                            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-stone-200 bg-stone-50 px-2 py-1">
                              <div className="text-[11px] text-stone-600">
                                {laneJobs.length} jobs - {totalHours.toFixed(2)} hrs
                              </div>
                              <div className="flex flex-wrap items-center gap-2">
                                <input
                                  type="text"
                                  value={operatorName}
                                  onChange={(event) => updatePressOperator(day.key, press, event.target.value)}
                                  onBlur={(event) => commitPressOperator(day.key, press, operatorName, event.target.value)}
                                  placeholder="Operator"
                                  disabled={!userCanEdit}
                                  className="min-w-[104px] border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800 outline-none focus:border-emerald-800 disabled:cursor-not-allowed disabled:bg-stone-100"
                                />
                                <input
                                  type="text"
                                  value={dutyName}
                                  onChange={(event) => updatePressDuty(day.key, press, event.target.value)}
                                  onBlur={(event) => commitPressDuty(day.key, press, dutyName, event.target.value)}
                                  placeholder="Duty"
                                  disabled={!userCanEdit}
                                  className="min-w-[104px] border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800 outline-none focus:border-emerald-800 disabled:cursor-not-allowed disabled:bg-stone-100"
                                />
                                <span className="border border-stone-300 bg-white px-2 py-1 text-[10px] font-medium text-stone-700">
                                  drop
                                </span>
                              </div>
                            </div>
                            <div className="space-y-0">
                              {laneJobs.map(({ assignment, job }, index) => (
                                <CompactScheduleCard
                                  key={assignment.id}
                                  job={job}
                                  assignment={assignment}
                                  finishMeta={job ? finishedMetaByJobId.get(job.id) : null}
                                  density={scheduleCardDensity}
                                  selected={job ? selectedJobId === job.id : false}
                                  onSelect={job ? () => selectJob(job.id) : undefined}
                                  canMoveUp={userCanMoveJobs && index > 0}
                                  canMoveDown={userCanMoveJobs && index < laneJobs.length - 1}
                                  onMoveUp={userCanMoveJobs && index > 0 ? () => moveAssignmentByStep(assignment.id, -1) : undefined}
                                  onMoveDown={userCanMoveJobs && index < laneJobs.length - 1 ? () => moveAssignmentByStep(assignment.id, 1) : undefined}
                                  onDropBefore={userCanMoveJobs ? (event) => handleScheduleCardDrop(event, day.key, press, assignment.id) : undefined}
                                  onUnschedule={userCanMoveJobs ? () => removeAssignment(assignment.id) : undefined}
                                  onFinish={job && userCanEdit ? () => finishJob(job.id) : undefined}
                                  onUndoFinish={job && userCanEdit && assignment.status === "finished" ? () => undoFinishJob(job.id) : undefined}
                                  onDuplicate={userCanMoveJobs ? () => duplicateAssignmentToNextDay(assignment.id) : undefined}
                                  onPickUp={userCanMoveJobs ? () => pickUpScheduledAssignment(assignment.id) : undefined}
                                  draggable={userCanMoveJobs}
                                />
                              ))}
                              {!laneJobs.length && (
                                <div className="min-h-16 border-dashed border-stone-200 p-3 text-center text-[11px] text-stone-500">
                                  Drop here
                                </div>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </React.Fragment>
                  ))}

                  {!visibleScheduleRows.length && (
                    <div className="col-span-6 rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No presses are visible for this week with the current filters.
                    </div>
                  )}
                </div>
              </div>
            </div>

            {!boardFocusMode && (
              <div className="grid gap-4 2xl:grid-cols-[minmax(0,1fr)_360px]">
                <div ref={jobDetailsRef} className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                  <div className="mb-3 flex items-center justify-between">
                    <div>
                      <div className="text-sm font-semibold">Job details</div>
                      <div className="text-xs text-stone-600">Select a job from the queue or board.</div>
                    </div>
                    {selectedJob && (
                      <span
                        className={`rounded-full px-2 py-1 text-xs font-medium ${statusTone(
                          deriveVisibleJobState(selectedJob.id, activePressJobIds, userFinishedJobIds)
                        )}`}
                      >
                        {deriveVisibleJobState(selectedJob.id, activePressJobIds, userFinishedJobIds)}
                      </span>
                    )}
                  </div>

                  {selectedJob ? (
                    <div className="space-y-3 text-sm">
                      <div>
                        <div className="text-lg font-semibold">
                          {selectedJob.customerName} {selectedJob.number}
                        </div>
                        <div className="text-stone-700">{selectedJob.generalDescr}</div>
                      </div>
                      <div className="grid grid-cols-2 gap-3 text-xs text-stone-700">
                        <Detail label="Default press" value={selectedJob.press ? formatPressLabel(selectedJob.press) : "-"} />
                        <Detail label="Priority" value={selectedJob.priority || "-"} />
                        <Detail label="Ship by" value={formatDate(selectedJob.shipByDate)} />
                        <Detail label="Imported status" value={selectedJob.ticketStatus || "-"} />
                        <Detail label="Quantity" value={selectedJob.ticQuantity.toLocaleString()} />
                        <Detail label="EST time" value={`${selectedJob.estPressTime.toFixed(2)} hrs`} />
                        <Detail label="PO number" value={selectedJob.custPoNum || "-"} />
                        <Detail label="Main tool" value={selectedJob.mainTool || "-"} />
                        <Detail label="Footage" value={selectedJob.estFootage.toLocaleString()} />
                        <Detail label="Stock" value={selectedJob.stockDisplay || "-"} />
                      </div>
                      {selectedJobFinishMeta && (
                        <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-3 text-xs text-emerald-950">
                          <div>
                            Marked done on {formatDateTime(selectedJobFinishMeta.finishedAt)}
                            {selectedJobFinishMeta.finishedBy ? ` by ${selectedJobFinishMeta.finishedBy}` : ""}
                          </div>
                          {userCanEdit && (
                            <button
                              type="button"
                              onClick={() => undoFinishJob(selectedJob.id)}
                              className="mt-3 rounded-xl border border-emerald-300 bg-white px-3 py-2 text-[11px] font-medium text-emerald-950"
                            >
                              Unmark done
                            </button>
                          )}
                        </div>
                      )}
                      <div className={`rounded-2xl border p-4 ${selectedJob.holdActive ? "border-rose-300 bg-rose-50" : "border-stone-300 bg-white/60"}`}>
                        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                          <div>
                            <div className="text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Job hold</div>
                            <div className="mt-1 text-sm text-stone-700">Keep a red hold note on this job even after new TXT imports.</div>
                          </div>
                          <label className="flex items-center gap-2 text-sm font-medium text-stone-800">
                            <input
                              type="checkbox"
                              checked={!!selectedJob.holdActive}
                              onChange={(event) => updateJobHoldState(selectedJob.id, event.target.checked)}
                              disabled={!userCanEdit}
                              className="h-4 w-4"
                            />
                            <span>Highlight this job as hold</span>
                          </label>
                        </div>
                        <div className="mt-3">
                          <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Hold reason</div>
                          <textarea
                            value={selectedJob.holdNote || ""}
                            onChange={(event) => updateJobHoldNote(selectedJob.id, event.target.value)}
                            disabled={!userCanEdit}
                            placeholder="Why is this job on hold?"
                            className="h-24 w-full rounded-2xl border border-stone-300 bg-white px-3 py-3 text-sm outline-none focus:border-emerald-800 disabled:cursor-not-allowed disabled:bg-stone-100"
                          />
                        </div>
                      </div>
                      <div>
                        <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Imported notes</div>
                        <div className="max-h-56 overflow-auto whitespace-pre-wrap rounded-2xl bg-stone-100 p-3 text-sm text-stone-800">
                          {selectedJob.notes || "No notes on this job."}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No job selected yet.
                    </div>
                  )}
                </div>

                <div className="space-y-4">
                  <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                    <div className="mb-3">
                      <div className="text-sm font-semibold">Job location search</div>
                      <div className="text-xs text-stone-600">Search any ticket, customer, or description to see every scheduled or finished day and press.</div>
                    </div>
                    <input
                      type="text"
                      value={locationSearch}
                      onChange={(event) => setLocationSearch(event.target.value)}
                      placeholder="Search job number, customer, or description"
                      className="w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                    />
                    <div className="mt-3 max-h-[36vh] space-y-3 overflow-y-auto">
                      {jobLocationResults.map(({ job, locations }) => (
                        <div key={job.id} className="rounded-2xl border border-stone-300 bg-white p-3">
                          <div className="flex items-start justify-between gap-3">
                            <div>
                              <div className="text-sm font-semibold">{job.customerName} {job.number}</div>
                              <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                            </div>
                            <button
                              onClick={() => selectJob(job.id, true)}
                              className="rounded-xl border border-stone-300 bg-stone-50 px-2 py-1 text-[11px] text-stone-800"
                            >
                              Open
                            </button>
                          </div>
                          <div className="mt-3 space-y-2">
                            {locations.length ? locations.map((location) => (
                              <button
                                key={location.id}
                                onClick={() => {
                                  setWeekStart(startOfWeek(new Date(`${location.dayKey}T12:00:00`)));
                                  selectJob(job.id, true);
                                }}
                                className="flex w-full items-center justify-between rounded-2xl bg-stone-100 px-3 py-2 text-left text-xs text-stone-800"
                              >
                                <span>{location.dayKey}</span>
                                <span>
                                  {formatPressLabel(location.press)} · {location.status === "finished" ? "done" : location.status}
                                </span>
                              </button>
                            )) : (
                              <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-3 text-xs text-stone-600">
                                This job is not on the schedule yet.
                              </div>
                            )}
                          </div>
                        </div>
                      ))}
                      {!jobLocationResults.length && (
                        <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                          {locationSearch.trim() ? "No jobs matched that search." : "Type a search to see where jobs are scheduled or finished."}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                    <div className="mb-3">
                      <div className="text-sm font-semibold">Ship ready today</div>
                      <div className="text-xs text-stone-600">These are only the jobs marked finished today by a logged-in user.</div>
                    </div>
                    <div className="max-h-[40vh] space-y-3 overflow-y-auto">
                      {doneJobs
                        .filter((job) => sameDay(job.finishMeta?.finishedAt, todayKey()))
                        .map((job) => (
                          <JobCard
                            key={job.id}
                            job={job}
                            state="ship"
                            selected={selectedJobId === job.id}
                            onClick={() => selectJob(job.id)}
                            finishedAt={job.finishMeta?.finishedAt}
                            finishedBy={job.finishMeta?.finishedBy}
                            onUndoFinish={userCanEdit ? () => undoFinishJob(job.id) : undefined}
                          />
                        ))}
                      {!doneJobs.filter((job) => sameDay(job.finishMeta?.finishedAt, todayKey())).length && (
                        <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                          No jobs marked finished today yet.
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {activeTab === "Notes" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">My internal checklist</div>
                <div className="text-xs text-stone-600">These notes are tied only to the signed-in user.</div>
              </div>
              <form onSubmit={addNote} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Checklist item</div>
                  <textarea
                    value={noteForm.text}
                    onChange={(event) => setNoteForm({ text: event.target.value })}
                    placeholder="Write a reminder, checklist item, or personal note"
                    className="h-36 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  type="submit"
                  disabled={!userCanEdit}
                  className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Add note
                </button>
              </form>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-sm font-semibold">My notes</div>
                  <div className="text-xs text-stone-600">Only {currentUser.username} can see these checklist items.</div>
                </div>
                <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">
                  {userNotes.filter((note) => !note.completed).length} open
                </div>
              </div>
              <div className="space-y-3">
                {userNotes.map((note) => (
                  <div key={note.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="flex gap-3">
                      <input
                        type="checkbox"
                        checked={note.completed}
                        disabled={!userCanEdit}
                        onChange={() => toggleNote(note.id)}
                        className="mt-1 h-4 w-4"
                      />
                      <div className="min-w-0 flex-1">
                        <div className={`whitespace-pre-wrap text-sm ${note.completed ? "text-stone-500 line-through" : "text-stone-800"}`}>
                          {note.text}
                        </div>
                        <div className="mt-2 text-xs text-stone-600">
                          Added {formatDateTime(note.createdAt)}
                          {note.completedAt ? ` · completed ${formatDateTime(note.completedAt)}` : ""}
                        </div>
                      </div>
                      <button
                        onClick={() => deleteNote(note.id)}
                        disabled={!userCanEdit}
                        className="rounded-xl border border-rose-200 px-3 py-2 text-xs text-rose-700 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Delete
                      </button>
                    </div>
                  </div>
                ))}
                {!userNotes.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                    No personal notes yet.
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "New Request" && (
          <div className="grid gap-4 xl:grid-cols-[minmax(0,760px)_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Request form</div>
                <div className="text-xs text-stone-600">Requests are stamped with the logged-in account, and you can attach supporting documents before saving.</div>
              </div>
              <form onSubmit={submitRequest} className="grid gap-4">
                <Field
                  label="Job number"
                  value={requestForm.jobNumber}
                  onChange={(value) => setRequestForm((current) => ({ ...current, jobNumber: value }))}
                  placeholder="Internal request number"
                />
                <Field
                  label="Customer"
                  value={requestForm.customer}
                  onChange={(value) => setRequestForm((current) => ({ ...current, customer: value }))}
                  placeholder="Customer name"
                />
                <Field
                  label="Requestor name"
                  value={requestForm.requestorName}
                  onChange={(value) => setRequestForm((current) => ({ ...current, requestorName: value }))}
                  placeholder="Who is asking for it"
                />
                <div className="grid gap-4 md:grid-cols-2">
                  <div>
                    <div className="mb-2 text-sm font-medium text-stone-800">Request type</div>
                    <select
                      value={requestForm.requestType}
                      onChange={(event) => setRequestForm((current) => ({ ...current, requestType: event.target.value }))}
                      className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                    >
                      <option>General Request</option>
                      <option>Ticket Checklist</option>
                    </select>
                  </div>
                  <div>
                    <div className="mb-2 text-sm font-medium text-stone-800">Send to one user</div>
                    <select
                      value={requestForm.assignedToAccount}
                      onChange={(event) => setRequestForm((current) => ({ ...current, assignedToAccount: event.target.value }))}
                      className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                    >
                      <option value="">Everyone with access</option>
                      {requestAssigneeOptions.map((username) => (
                        <option key={username} value={username}>
                          {username}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Description</div>
                  <textarea
                    value={requestForm.description}
                    onChange={(event) =>
                      setRequestForm((current) => ({ ...current, description: event.target.value }))
                    }
                    placeholder="Include QTY and any extra details here"
                    className="h-36 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                  <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                    <div className="font-medium text-stone-900">Upload PDF, Word, Excel, or other request files</div>
                    <input
                      type="file"
                      accept={ATTACHMENT_ACCEPT}
                      multiple
                      onChange={handleDraftAttachmentChange}
                      className="mt-3 block w-full text-xs"
                    />
                  </label>
                </div>
                <AttachmentList attachments={requestDraftAttachments} onRemove={removeDraftAttachment} />
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save request
                  </button>
                  <button
                    type="button"
                    onClick={clearRequestDraft}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-4">
                <div className="text-sm font-semibold">Request workflow</div>
                <div className="text-xs text-stone-600">Use Ticket Checklist when a job needs a second check, and optionally route it to just one user instead of the full queue.</div>
              </div>
              <div className="grid gap-3">
                <StatRow label="Open requests" value={openRequests.length} />
                <StatRow label="Completed requests" value={requestHistory.length} />
                <StatRow label="Signed-in account" value={currentUser.username} />
                {canAccessTab(currentUser, "Open Requests") && (
                  <button
                    onClick={() => setActiveTab("Open Requests")}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-left text-sm text-stone-800"
                  >
                    Open the request queue
                  </button>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "Open Requests" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Open requests</div>
                <div className="text-xs text-stone-600">Mark done to move a request into history under your login. Uploads added here stay attached to the request.</div>
              </div>
              <button
                onClick={() => setActiveTab("New Request")}
                className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
              >
                New request
              </button>
            </div>

            <div className="space-y-3">
                  {openRequests.map((request) => (
                    <RequestCard
                      key={request.id}
                      request={request}
                      onDone={() => markRequestDone(request.id)}
                      onDelete={() => deleteRequest(request.id)}
                      onAddAttachments={(files) => addRequestAttachments(request.id, files)}
                      onRemoveAttachment={(attachmentId) => removeRequestAttachment(request.id, attachmentId)}
                      readOnly={!userCanEdit}
                    />
                  ))}
              {!openRequests.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No open requests right now.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "Request History" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Request history</div>
                <div className="text-xs text-stone-600">Completed requests are sorted by completion time and show which login completed them.</div>
              </div>
              <div className="flex items-end gap-2">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                  <input
                    type="date"
                    value={requestHistoryFilterDate}
                    onChange={(event) => setRequestHistoryFilterDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  onClick={() => setRequestHistoryFilterDate("")}
                  className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                >
                  Clear
                </button>
              </div>
            </div>
            <div className="space-y-3">
              {requestHistory.map((request) => (
                <RequestCard key={request.id} request={request} readOnly />
              ))}
              {!requestHistory.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No completed requests yet.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "Pull Paper Request" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Pull paper request</div>
                <div className="text-xs text-stone-600">Send a paper pull note to a press or digital and keep the open list here.</div>
              </div>
              <form onSubmit={submitPullPaperRequest} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Send to</div>
                  <select
                    value={pullPaperForm.target}
                    onChange={(event) => setPullPaperForm((current) => ({ ...current, target: event.target.value }))}
                    className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  >
                    {PULL_PAPER_TARGETS.map((target) => (
                      <option key={target} value={target}>
                        {target}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Paper note</div>
                  <textarea
                    value={pullPaperForm.details}
                    onChange={(event) => setPullPaperForm((current) => ({ ...current, details: event.target.value }))}
                    placeholder='Stock 266 Width 9.5" 1 roll'
                    className="h-40 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                  Requesting as <span className="font-semibold">{currentUser.username}</span>
                </div>
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save pull request
                  </button>
                  <button
                    type="button"
                    onClick={() => setPullPaperForm(EMPTY_PULL_PAPER_FORM)}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Open pull paper requests</div>
                    <div className="text-xs text-stone-600">Mark done when the paper is pulled, or delete it if it is no longer needed.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{openPullPaperRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {openPullPaperRequests.map((request) => (
                    <PaperPullCard
                      key={request.id}
                      request={request}
                      onDone={() => markPullPaperRequestDone(request.id)}
                      onDelete={() => deletePullPaperRequest(request.id)}
                      readOnly={!userCanEdit}
                    />
                  ))}
                  {!openPullPaperRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No open pull paper requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Completed pull paper requests</div>
                    <div className="text-xs text-stone-600">Recent completed paper pulls stay here so you can see who closed them out.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{completedPullPaperRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {completedPullPaperRequests.slice(0, 12).map((request) => (
                    <PaperPullCard key={request.id} request={request} readOnly />
                  ))}
                  {!completedPullPaperRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No completed paper pull requests yet.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "Supplies Request" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Supplies request</div>
                <div className="text-xs text-stone-600">Create a supplies request, attach supporting files if needed, and keep the open count visible in the tab.</div>
              </div>
              <form onSubmit={submitSuppliesRequest} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Supplies note</div>
                  <textarea
                    value={suppliesForm.details}
                    onChange={(event) => setSuppliesForm((current) => ({ ...current, details: event.target.value }))}
                    placeholder="Boxes, labels, cores, bags, tape, or any other supply request details"
                    className="h-40 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                  <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                    <div className="font-medium text-stone-900">Upload PDF, PNG, JPG, JPEG, or other supply request files</div>
                    <input
                      type="file"
                      accept={ATTACHMENT_ACCEPT}
                      multiple
                      onChange={handleSuppliesDraftAttachmentChange}
                      className="mt-3 block w-full text-xs"
                    />
                  </label>
                </div>
                <AttachmentList attachments={suppliesDraftAttachments} onRemove={removeSuppliesDraftAttachment} />
                <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                  Requesting as <span className="font-semibold">{currentUser.username}</span>
                </div>
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save supplies request
                  </button>
                  <button
                    type="button"
                    onClick={clearSuppliesDraft}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Open supplies requests</div>
                    <div className="text-xs text-stone-600">Mark done when the supplies request is complete, or delete it if it is no longer needed.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{openSuppliesRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {openSuppliesRequests.map((request) => (
                    <SuppliesRequestCard
                      key={request.id}
                      request={request}
                      onDone={() => markSuppliesRequestDone(request.id)}
                      onDelete={() => deleteSuppliesRequest(request.id)}
                      onAddAttachments={(files) => addSuppliesRequestAttachments(request.id, files)}
                      onRemoveAttachment={(attachmentId) => removeSuppliesRequestAttachment(request.id, attachmentId)}
                      readOnly={!userCanEdit}
                    />
                  ))}
                  {!openSuppliesRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No open supplies requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Completed supplies requests</div>
                    <div className="text-xs text-stone-600">Recent completed supplies requests stay here so you can see who closed them out.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{completedSuppliesRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {completedSuppliesRequests.slice(0, 12).map((request) => (
                    <SuppliesRequestCard key={request.id} request={request} readOnly />
                  ))}
                  {!completedSuppliesRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No completed supplies requests yet.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "Daily Shipment" && (
          <div className="space-y-4">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-sm font-semibold">Daily shipment</div>
                  <div className="text-xs text-stone-600">
                    Finished jobs can be assigned to a ship date first, then grouped under that shipping day.
                  </div>
                  <div className="mt-2 text-xs text-stone-700">
                    {shipmentEmailsForSelectedDate.length
                      ? `${shipmentEmailsForSelectedDate.length} shipment email${shipmentEmailsForSelectedDate.length === 1 ? "" : "s"} already logged for this date.`
                      : "No shipment email has been logged for this date yet."}
                  </div>
                </div>
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Ship date</div>
                  <input
                    type="date"
                    value={selectedShipDate}
                    onChange={(event) => setSelectedShipDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-sm font-semibold">Assign ship date</div>
                  <div className="text-xs text-stone-600">
                    Pick the finished jobs that have not been grouped yet, assign them to a ship date, and they will move into that day's shipment queue.
                  </div>
                  <div className="mt-1 text-[11px] text-stone-600">
                    Showing {SHIPMENT_QUEUE_WINDOW_OPTIONS.find((option) => option.value === shipQueueWindow)?.label.toLowerCase() || "recent finished jobs"} so this tab stays fast.
                  </div>
                </div>
                <div className="flex flex-wrap items-end gap-2">
                  <div>
                    <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Finished within</div>
                    <select
                      value={shipQueueWindow}
                      onChange={(event) => setShipQueueWindow(event.target.value)}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                    >
                      {SHIPMENT_QUEUE_WINDOW_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Ship date</div>
                    <input
                      type="date"
                      value={shipDateDraft}
                      onChange={(event) => setShipDateDraft(event.target.value)}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => setSelectedShipQueueJobs(unassignedFinishedJobs.map((job) => job.id))}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    Select all
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedShipQueueJobs([])}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    Clear
                  </button>
                  <button
                    type="button"
                    onClick={assignShipDateToFinishedJobs}
                    className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
                  >
                    Assign date
                  </button>
                </div>
              </div>

              <div className="mb-3 flex items-center justify-between">
                <div className="text-xs text-stone-600">
                  {selectedShipQueueJobs.length} selected
                </div>
                <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{unassignedFinishedJobs.length}</div>
              </div>

                <div className="max-h-[320px] space-y-3 overflow-y-auto pr-1">
                  {unassignedFinishedJobs.map((job) => (
                    <div key={job.id} className="rounded-2xl border border-stone-300 bg-white p-3">
                      <div className="flex gap-3">
                        <input
                          type="checkbox"
                          checked={selectedShipQueueJobs.includes(job.id)}
                          onChange={() => toggleShipQueueJob(job.id)}
                          className="mt-1 h-4 w-4"
                        />
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-start justify-between gap-2">
                            <div>
                              <div className="text-sm font-semibold">
                                {job.customerName} {job.number}
                              </div>
                              <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                            </div>
                            <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone("done")}`}>done</span>
                          </div>
                          <div className="mt-2 grid grid-cols-2 gap-2 text-xs text-stone-700 md:grid-cols-4">
                            <InfoPill label="Press" value={job.press || "-"} />
                            <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
                            <InfoPill label="Finished" value={formatDateTime(job.finishMeta?.finishedAt)} />
                            <InfoPill label="Ship date" value={effectiveFinishedShipDate(job.finishMeta) || "-"} />
                          </div>
                          <div className="mt-3 flex flex-wrap gap-2">
                            <button
                              type="button"
                              onClick={() => dismissFinishedJobFromShippingQueue(job.id)}
                              className="rounded-xl border border-rose-200 px-3 py-2 text-xs text-rose-700"
                            >
                              Remove from queue
                            </button>
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                {!unassignedFinishedJobs.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                    All finished jobs are already grouped into shipments.
                  </div>
                )}
              </div>
            </div>

            <div className="grid gap-4 xl:grid-cols-[minmax(0,1.1fr)_420px]">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                  <div>
                    <div className="text-sm font-semibold">Date done on {selectedShipDate}</div>
                    <div className="text-xs text-stone-600">This queue follows the imported DateDone value from the TXT. Grouped or removed jobs stay visible so you can still track what was done.</div>
                    {!!dateDoneJobs.length && (
                      <div className="mt-2 text-[11px] text-stone-700">
                        Jobs in queue: {dateDoneJobNumbers}
                      </div>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setSelectedShipmentJobs(readyToShipJobs.map((job) => job.id))}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                    >
                      Select all
                    </button>
                    <button
                      type="button"
                      onClick={() => setSelectedShipmentJobs([])}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                    >
                      Clear
                    </button>
                    <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{dateDoneJobs.length}</div>
                  </div>
                </div>

                <div className="space-y-3">
                  {dateDoneJobs.map((job) => {
                    const isGrouped = assignedShipmentJobIds.has(job.id);
                    const isRemoved = Boolean(job.finishMeta?.excludeFromShipping);
                    const isSelectable = !isGrouped && !isRemoved;
                    const statusLabel = isGrouped ? "grouped" : isRemoved ? "removed" : "ready";
                    return (
                    <label
                      key={job.id}
                      className={`flex gap-3 rounded-2xl border p-3 ${
                        isSelectable ? "border-stone-300 bg-white" : "border-stone-200 bg-stone-100/80"
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={selectedShipmentJobs.includes(job.id)}
                        onChange={() => toggleShipmentJob(job.id)}
                        disabled={!isSelectable}
                        className="mt-1 h-4 w-4"
                      />
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-start justify-between gap-2">
                          <div>
                            <div className="text-sm font-semibold">
                              {job.customerName} {job.number}
                            </div>
                            <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                          </div>
                          <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(isSelectable ? "done" : "blocked")}`}>{statusLabel}</span>
                        </div>
                        <div className="mt-2 grid grid-cols-2 gap-2 text-xs text-stone-700 md:grid-cols-4">
                          <InfoPill label="Press" value={job.press || "-"} />
                          <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
                          <InfoPill label="Finished" value={formatDateTime(job.finishMeta?.finishedAt)} />
                          <InfoPill label="By" value={job.finishMeta?.finishedBy || "-"} />
                        </div>
                        {!isSelectable && (
                          <div className="mt-2 text-xs text-stone-600">
                            {isGrouped ? "Already included in a shipment group." : "Removed from the shipping queue."}
                          </div>
                        )}
                      </div>
                    </label>
                  )})}
                  {!dateDoneJobs.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No finished jobs match that done date.
                    </div>
                  )}
                </div>
              </div>

              <div className="space-y-4">
                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Create shipment group</div>
                    <div className="text-xs text-stone-600">Example: one skid with 3 jobs for $141.17, separate FedEx transactions, or a custom method like JP Express.</div>
                  </div>
                  <div className="mb-4 rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="mb-3 text-sm font-semibold">Shipping methods</div>
                    <div className="mb-3 flex gap-2">
                      <input
                        type="text"
                        value={newShipmentMethod}
                        onChange={(event) => setNewShipmentMethod(event.target.value)}
                        placeholder="Add a method like JP Express"
                        className="flex-1 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      />
                      <button
                        type="button"
                        onClick={addShipmentMethod}
                        className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white"
                      >
                        Add
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {shipmentMethods.map((method) => (
                        <div key={method} className="flex items-center gap-2 rounded-full border border-stone-300 bg-stone-100 px-3 py-2 text-xs text-stone-800">
                          <span>{method}</span>
                          <button
                            type="button"
                            onClick={() => removeShipmentMethod(method)}
                            className="text-rose-700"
                          >
                            remove
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                  <form onSubmit={createShipmentGroup} className="grid gap-3">
                    <Field
                      label="Label"
                      value={shipmentForm.label}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, label: value }))}
                      placeholder="Skid A or FedEx 1"
                    />
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Method</div>
                      <select
                        value={shipmentForm.method}
                        onChange={(event) => setShipmentForm((current) => ({ ...current, method: event.target.value }))}
                        className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      >
                        {shipmentMethods.map((method) => (
                          <option key={method} value={method}>
                            {method}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="grid grid-cols-[minmax(0,1fr)_160px] gap-3">
                      <Field
                        label="Skids / cartons out"
                        value={shipmentForm.packageCount}
                        onChange={(value) => setShipmentForm((current) => ({ ...current, packageCount: value }))}
                        placeholder="3"
                      />
                      <div>
                        <div className="mb-2 text-sm font-medium text-stone-800">Type</div>
                        <select
                          value={shipmentForm.packageType}
                          onChange={(event) =>
                            setShipmentForm((current) => ({ ...current, packageType: event.target.value }))
                          }
                          className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                        >
                          <option>Skids</option>
                          <option>Cartons</option>
                        </select>
                      </div>
                    </div>
                    <Field
                      label="Total price"
                      value={shipmentForm.totalCost}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, totalCost: value }))}
                      placeholder="141.17"
                    />
                    <Field
                      label="Bill customer"
                      value={shipmentForm.billAmount}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, billAmount: value }))}
                      placeholder="185.00"
                    />
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Notes</div>
                      <textarea
                        value={shipmentForm.notes}
                        onChange={(event) => setShipmentForm((current) => ({ ...current, notes: event.target.value }))}
                        placeholder="Optional shipment notes"
                        className="h-24 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      />
                    </div>
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                      <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                        <div className="font-medium text-stone-900">Upload PDF, PNG, JPG, JPEG, or other shipment files</div>
                        <input
                          type="file"
                          accept={ATTACHMENT_ACCEPT}
                          multiple
                          onChange={handleShipmentDraftAttachmentChange}
                          className="mt-3 block w-full text-xs"
                        />
                      </label>
                    </div>
                    <AttachmentList attachments={shipmentDraftAttachments} onRemove={removeShipmentDraftAttachment} />
                    <div className="flex flex-wrap gap-2">
                      <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                        Create shipment group
                      </button>
                      <button
                        type="button"
                        onClick={clearShipmentDraft}
                        className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                      >
                        Reset draft
                      </button>
                    </div>
                  </form>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Daily shipment email</div>
                    <div className="text-xs text-stone-600">Open a draft email for this ship date with normal spacing, then log it so everyone can see it was already sent.</div>
                  </div>
                  <div className="grid gap-3">
                    <Field
                      label="Recipients"
                      value={shipmentEmailForm.recipients}
                      onChange={(value) => setShipmentEmailForm((current) => ({ ...current, recipients: value }))}
                      placeholder="shipping@company.com; billing@company.com"
                    />
                    <Field
                      label="CC"
                      value={shipmentEmailForm.cc}
                      onChange={(value) => setShipmentEmailForm((current) => ({ ...current, cc: value }))}
                      placeholder="manager@company.com; office@company.com"
                    />
                    <div className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="mb-3 text-sm font-semibold">Email groups</div>
                      <div className="grid gap-3">
                        <div className="flex flex-col gap-2 md:flex-row">
                          <select
                            value={shipmentEmailForm.groupId}
                            onChange={(event) => applyShipmentEmailGroup(event.target.value)}
                            className="flex-1 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                          >
                            <option value="">Select a saved email group</option>
                            {shipmentEmailGroups.map((group) => (
                              <option key={group.id} value={group.id}>
                                {group.name}
                              </option>
                            ))}
                          </select>
                          <button
                            type="button"
                            disabled={!userCanEdit || !shipmentEmailForm.groupId}
                            onClick={() => removeShipmentEmailGroup(shipmentEmailForm.groupId)}
                            className="rounded-2xl border border-rose-200 px-4 py-3 text-sm text-rose-700 disabled:cursor-not-allowed disabled:opacity-50"
                          >
                            Remove group
                          </button>
                        </div>
                        <Field
                          label="Group name"
                          value={shipmentEmailGroupForm.name}
                          onChange={(value) => setShipmentEmailGroupForm((current) => ({ ...current, name: value }))}
                          placeholder="Daily shipment list"
                        />
                        <Field
                          label="Group recipients"
                          value={shipmentEmailGroupForm.recipients}
                          onChange={(value) => setShipmentEmailGroupForm((current) => ({ ...current, recipients: value }))}
                          placeholder="team1@company.com; team2@company.com"
                        />
                        <Field
                          label="Group CC"
                          value={shipmentEmailGroupForm.cc}
                          onChange={(value) => setShipmentEmailGroupForm((current) => ({ ...current, cc: value }))}
                          placeholder="lead@company.com; accounting@company.com"
                        />
                        <button
                          type="button"
                          disabled={!userCanEdit}
                          onClick={saveShipmentEmailGroup}
                          className="rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800 disabled:cursor-not-allowed disabled:opacity-50"
                        >
                          Save group
                        </button>
                      </div>
                    </div>
                    <div className="rounded-2xl bg-stone-100 p-4 text-sm text-stone-800">
                      <div className="font-semibold">{shipmentEmailDraft.subject}</div>
                      <div className="mt-2 text-xs text-stone-600">To: {shipmentEmailForm.recipients || "-"}</div>
                      <div className="mt-1 text-xs text-stone-600">CC: {shipmentEmailForm.cc || "-"}</div>
                      <pre className="mt-2 max-h-52 overflow-auto whitespace-pre-wrap text-xs text-stone-700">{shipmentEmailDraft.body}</pre>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        type="button"
                        disabled={!userCanEdit}
                        onClick={openShipmentEmailDraft}
                        className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Open email draft
                      </button>
                      <button
                        type="button"
                        disabled={!userCanEdit}
                        onClick={logShipmentEmail}
                        className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Mark emailed
                      </button>
                    </div>
                  </div>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Shipment history</div>
                    <div className="text-xs text-stone-600">Pick a past date to see the exact jobs that shipped that day.</div>
                  </div>
                  <div className="mb-4 flex items-end gap-2">
                    <div className="flex-1">
                      <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                      <input
                        type="date"
                        value={shipmentHistoryFilterDate}
                        onChange={(event) => setShipmentHistoryFilterDate(event.target.value)}
                        className="w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                      />
                    </div>
                    <button
                      onClick={() => setShipmentHistoryFilterDate("")}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                    >
                      Clear
                    </button>
                  </div>
                  <div className="space-y-3">
                    {shipmentHistoryDays.map((item) => (
                      <button
                        key={item.shipDate}
                        onClick={() => setSelectedShipDate(item.shipDate)}
                        className={`w-full rounded-2xl border px-4 py-3 text-left ${
                          selectedShipDate === item.shipDate
                            ? "border-emerald-900 bg-emerald-900 text-white"
                            : "border-stone-300 bg-white text-stone-800"
                        }`}
                      >
                        <div className="text-sm font-semibold">{item.shipDate}</div>
                        <div className={`mt-1 text-xs ${selectedShipDate === item.shipDate ? "text-stone-200" : "text-stone-600"}`}>
                          {item.groupCount} groups - {item.jobCount} jobs - {formatCurrency(item.totalCost)}
                        </div>
                      </button>
                    ))}
                    {!shipmentHistoryDays.length && (
                      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                        No shipment history yet.
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="mb-4">
                <div className="text-sm font-semibold">Shipped on {selectedShipDate}</div>
                <div className="text-xs text-stone-600">Each shipment group keeps the exact jobs and who marked them finished.</div>
              </div>
              <div className="space-y-3">
                {shipmentGroupsForDay.map((group) => {
                  const items = shipmentItemsByGroupId.get(group.id) || [];
                  return (
                    <div key={group.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                        <div>
                          <div className="flex flex-wrap items-center gap-2">
                            <div className="text-sm font-semibold">{group.label}</div>
                            <span className="rounded-full bg-emerald-900 px-2 py-1 text-[11px] font-medium text-white">
                              {group.method}
                            </span>
                          </div>
                          <div className="mt-1 text-xs text-stone-600">
                            {items.length} jobs - {formatCurrency(group.totalCost)} - created {formatDateTime(group.createdAt)}
                            {group.createdBy ? ` by ${group.createdBy}` : ""}
                          </div>
                          <div className="mt-1 text-xs text-stone-600">
                            Bill {formatCurrency(group.billAmount)}
                          </div>
                          {!!group.packageCount && (
                            <div className="mt-1 text-xs text-stone-600">
                              {group.packageCount} {safeText(group.packageType || "Skids").toLowerCase()}
                            </div>
                          )}
                          {group.notes && <div className="mt-2 text-sm text-stone-800">{group.notes}</div>}
                          <div className="mt-3">
                            <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
                            <AttachmentList
                              attachments={group.attachments || []}
                              onRemove={(attachmentId) => removeShipmentGroupAttachment(group.id, attachmentId)}
                            />
                          </div>
                          <label className="mt-3 block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
                            <div className="font-medium text-stone-900">Add more shipment files</div>
                            <input
                              type="file"
                              accept={ATTACHMENT_ACCEPT}
                              multiple
                              onChange={(event) => {
                                const files = event.target.files;
                                if (files?.length) addShipmentGroupAttachments(group.id, files);
                                event.target.value = "";
                              }}
                              className="mt-3 block w-full text-xs"
                            />
                          </label>
                        </div>
                        <button
                          onClick={() => deleteShipmentGroup(group.id)}
                          className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700"
                        >
                          Delete group
                        </button>
                      </div>
                      <div className="mt-3 grid gap-2 md:grid-cols-2">
                        {items.map((item) => (
                            <div key={item.id} className="rounded-2xl bg-stone-100 p-3">
                            <div className="text-sm font-semibold">
                              {item.customerName} {item.number}
                            </div>
                              <div className="mt-1 text-xs text-stone-700">{item.generalDescr}</div>
                              <div className="mt-2 text-[11px] text-stone-600">
                              Finished {formatDateTime(item.finishedAt)} by {item.finishedBy || "-"}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
                {!shipmentGroupsForDay.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                    No shipment groups created for this date yet.
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "Shipment Emails" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Shipment email log</div>
                <div className="text-xs text-stone-600">Use this to confirm which shipment dates were already emailed so you do not double-send them.</div>
              </div>
              <div className="flex items-end gap-2">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                  <input
                    type="date"
                    value={shipmentEmailHistoryFilterDate}
                    onChange={(event) => setShipmentEmailHistoryFilterDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  onClick={() => setShipmentEmailHistoryFilterDate("")}
                  className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                >
                  Clear
                </button>
              </div>
            </div>
            <div className="space-y-3">
              {shipmentEmailHistory.map((log) => (
                <div key={log.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                  <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <div className="text-sm font-semibold">{log.shipDate}</div>
                        <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">
                          {log.jobCount} jobs
                        </span>
                      </div>
                      <div className="mt-1 text-xs text-stone-600">
                        Sent by {log.createdBy || "-"} on {formatDateTime(log.createdAt)}
                      </div>
                      <div className="mt-1 text-xs text-stone-600">Recipients: {log.recipients || "-"}</div>
                      <div className="mt-1 text-xs text-stone-600">CC: {log.cc || "-"}</div>
                      <div className="mt-1 text-xs text-stone-600">
                        Methods: {log.methods.length ? log.methods.join(", ") : "-"} · Cost {formatCurrency(log.totalCost)} · Bill {formatCurrency(log.totalBill)}
                      </div>
                      <div className="mt-3 rounded-2xl bg-stone-100 p-3">
                        <div className="text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Subject</div>
                        <div className="mt-1 text-sm text-stone-800">{log.subject}</div>
                        <div className="mt-3 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Body</div>
                        <pre className="mt-1 whitespace-pre-wrap text-xs text-stone-700">{log.body}</pre>
                      </div>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        onClick={() => {
                          setSelectedShipDate(log.shipDate);
                          setActiveTab("Daily Shipment");
                        }}
                        className="rounded-2xl border border-stone-300 bg-stone-50 px-3 py-2 text-sm text-stone-800"
                      >
                        Open date
                      </button>
                      <button
                        onClick={() => deleteShipmentEmailLog(log.id)}
                        className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700"
                      >
                        Delete
                      </button>
                    </div>
                  </div>
                </div>
              ))}
              {!shipmentEmailHistory.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No shipment emails have been logged yet.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "Activity Log" && (
          <div className="space-y-4">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-4 flex flex-col gap-3 xl:flex-row xl:items-end xl:justify-between">
                <div>
                  <div className="text-sm font-semibold">Activity log</div>
                  <div className="text-xs text-stone-600">Track schedule moves, shipment updates, request changes, and user-permission changes.</div>
                </div>
                <div className="grid gap-2 sm:grid-cols-2 xl:grid-cols-4">
                  <select
                    value={activityFilters.actor}
                    onChange={(event) => setActivityFilters((current) => ({ ...current, actor: event.target.value }))}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    {activityActors.map((actor) => (
                      <option key={actor} value={actor}>
                        {actor === "All" ? "All people" : actor}
                      </option>
                    ))}
                  </select>
                  <select
                    value={activityFilters.scope}
                    onChange={(event) => setActivityFilters((current) => ({ ...current, scope: event.target.value }))}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    {activityScopes.map((scope) => (
                      <option key={scope} value={scope}>
                        {scope === "All" ? "All areas" : scope}
                      </option>
                    ))}
                  </select>
                  <input
                    type="date"
                    value={activityFilters.date}
                    onChange={(event) => setActivityFilters((current) => ({ ...current, date: event.target.value }))}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                  <button
                    onClick={() => setActivityFilters(EMPTY_ACTIVITY_FILTERS)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    Clear filters
                  </button>
                </div>
              </div>

              <div className="grid gap-3 md:grid-cols-3">
                <StatRow label="Entries shown" value={filteredActivityLog.length} />
                <StatRow label="People" value={Math.max(activityActors.length - 1, 0)} />
                <StatRow label="Areas" value={Math.max(activityScopes.length - 1, 0)} />
              </div>
            </div>

            <div className="space-y-3">
              {filteredActivityLog.map((entry) => {
                const detailEntries = Object.entries(entry.details || {}).filter(([, value]) =>
                  value !== null && value !== undefined && String(value).trim() !== ""
                );
                return (
                  <div key={entry.id} className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                    <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <div className="text-sm font-semibold text-stone-900">{entry.action}</div>
                          <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">
                            {entry.scope}
                          </span>
                          <span className="rounded-full bg-white px-2 py-1 text-[11px] text-stone-700">
                            {entry.actor}
                          </span>
                        </div>
                        <div className="mt-2 text-sm text-stone-800">{entry.description}</div>
                        {!!detailEntries.length && (
                          <div className="mt-3 flex flex-wrap gap-2">
                            {detailEntries.slice(0, 6).map(([key, value]) => (
                              <span key={key} className="rounded-full border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-700">
                                {key}: {String(value)}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                      <div className="rounded-2xl bg-white px-3 py-2 text-xs text-stone-600">
                        {formatDateTime(entry.createdAt)}
                      </div>
                    </div>
                  </div>
                );
              })}
              {!filteredActivityLog.length && (
                <div className="rounded-3xl border border-dashed border-stone-300 bg-white/60 p-6 text-sm text-stone-600">
                  No activity matched those filters yet.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "User Admin" && userCanManageUsers && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Add user</div>
                <div className="text-xs text-stone-600">Management can create logins, choose edit or view-only access, and decide which tabs each user can see.</div>
              </div>
              <form onSubmit={createUser} className="grid gap-4">
                <Field
                  label="Username"
                  value={userForm.username}
                  onChange={(value) => setUserForm((current) => ({ ...current, username: value }))}
                  placeholder="New username"
                />
                <Field
                  label="Password"
                  value={userForm.password}
                  onChange={(value) => setUserForm((current) => ({ ...current, password: value }))}
                  placeholder="Set a password"
                />
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Access</div>
                  <select
                    value={userForm.accessMode}
                    onChange={(event) => setUserForm((current) => ({ ...current, accessMode: event.target.value }))}
                    className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  >
                    {ACCESS_MODE_OPTIONS.map((mode) => (
                      <option key={mode} value={mode}>
                        {mode === "edit" ? "Edit" : "View only"}
                      </option>
                    ))}
                  </select>
                </div>
                <label className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800">
                  <input
                    type="checkbox"
                    checked={userForm.canManageUsers}
                    onChange={(event) => setUserForm((current) => ({ ...current, canManageUsers: event.target.checked }))}
                    className="h-4 w-4"
                  />
                  <span>Allow user admin access</span>
                </label>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Visible tabs</div>
                  <div className="grid gap-2 md:grid-cols-2">
                    {[...BASE_TABS, "User Admin"].map((tab) => (
                      <label key={tab} className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800">
                        <input
                          type="checkbox"
                          checked={userForm.canManageUsers && tab === "User Admin" ? true : userForm.tabs.includes(tab)}
                          disabled={tab === "User Admin"}
                          onChange={() =>
                            setUserForm((current) => {
                              const nextTabs = current.tabs.includes(tab)
                                ? current.tabs.filter((value) => value !== tab)
                                : [...current.tabs, tab];
                              return { ...current, tabs: normalizeUserTabs(nextTabs, "Warehouse/Shipper", false, current.canManageUsers) };
                            })
                          }
                          className="h-4 w-4"
                        />
                        <span>{tab}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                  Create user
                </button>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-5 flex items-center justify-between gap-3">
                  <div>
                    <div className="text-sm font-semibold">Registration approvals</div>
                    <div className="text-xs text-stone-600">Approve new account requests before they can log in.</div>
                  </div>
                  <div className="rounded-full bg-stone-200 px-3 py-1 text-xs font-semibold text-stone-800">
                    {pendingRegistrationRequests.length} pending
                  </div>
                </div>
                <div className="space-y-3">
                  {pendingRegistrationRequests.map((request) => (
                    <div key={request.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                        <div>
                          <div className="text-sm font-semibold">{request.username}</div>
                          <div className="mt-1 text-xs text-stone-600">
                            Requested {formatDateTime(request.createdAt)}
                          </div>
                        </div>
                        <div className="flex flex-wrap gap-2">
                          <button
                            onClick={() => approveRegistrationRequest(request.id)}
                            className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
                          >
                            Approve
                          </button>
                          <button
                            onClick={() => denyRegistrationRequest(request.id)}
                            className="rounded-2xl border border-stone-300 bg-white px-4 py-2 text-sm text-stone-800"
                          >
                            Deny
                          </button>
                        </div>
                      </div>
                    </div>
                  ))}
                  {!pendingRegistrationRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No pending registration requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-5">
                  <div className="text-sm font-semibold">Manage users</div>
                  <div className="text-xs text-stone-600">Rename users, reset passwords, choose edit or view-only, delete accounts, and control which tabs each account can see.</div>
                </div>
                <div className="space-y-3">
                  {users.map((user) => (
                    <div key={user.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="space-y-4">
                        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                          <div>
                            <div className="flex flex-wrap items-center gap-2">
                              <div className="text-sm font-semibold">{user.username}</div>
                              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${hasManagementAccess(user) ? "bg-emerald-900 text-white" : "bg-stone-200 text-stone-800"}`}>
                                {hasManagementAccess(user) ? "User Admin" : user.accessMode === "edit" ? "Edit" : "View only"}
                              </span>
                            </div>
                            <div className="mt-1 text-xs text-stone-600">
                              Created {formatDateTime(user.createdAt)} by {user.createdBy || "-"}
                            </div>
                          </div>
                          <div className="grid gap-2 md:w-[380px]">
                            <div className="grid gap-2 md:grid-cols-[1fr_auto]">
                              <input
                                type="text"
                                value={userUsernameDrafts[user.id] ?? user.username}
                                onChange={(event) =>
                                  setUserUsernameDrafts((current) => ({ ...current, [user.id]: event.target.value }))
                                }
                                placeholder={`Rename ${user.username}`}
                                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                              />
                              <button
                                onClick={() => renameUser(user.id)}
                                className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                              >
                                Save name
                              </button>
                            </div>
                            <div className="grid gap-2 md:grid-cols-[1fr_auto]">
                              <input
                                type="text"
                                value={userPasswordDrafts[user.id] || ""}
                                onChange={(event) =>
                                  setUserPasswordDrafts((current) => ({ ...current, [user.id]: event.target.value }))
                                }
                                placeholder={`Set new password for ${user.username}`}
                                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                              />
                              <button
                                onClick={() => updateUserPassword(user.id)}
                                className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                              >
                                Save password
                              </button>
                            </div>
                            <button
                              onClick={() => deleteUser(user.id)}
                              disabled={comparableUsername(user.username) === comparableUsername(currentUsername)}
                              className="rounded-2xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700 disabled:cursor-not-allowed disabled:opacity-50"
                            >
                              Delete user
                            </button>
                          </div>
                        </div>

                        <div className="grid gap-3 md:grid-cols-[180px_1fr]">
                          <div>
                            <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Access</div>
                            <select
                              value={user.accessMode}
                              onChange={(event) => updateUserAccess(user.id, { accessMode: event.target.value })}
                              className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                            >
                              {ACCESS_MODE_OPTIONS.map((mode) => (
                                <option key={mode} value={mode}>
                                  {mode === "edit" ? "Edit" : "View only"}
                                </option>
                              ))}
                            </select>
                          </div>
                          <label className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800">
                            <input
                              type="checkbox"
                              checked={!!user.canManageUsers}
                              onChange={(event) => updateUserAccess(user.id, { canManageUsers: event.target.checked })}
                              className="h-4 w-4"
                            />
                            <span>Can manage users and registration approvals</span>
                          </label>
                        </div>

                        <div>
                          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Visible tabs</div>
                          <div className="grid gap-2 md:grid-cols-2">
                            {[...BASE_TABS, "User Admin"].map((tab) => (
                              <label key={tab} className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800">
                                <input
                                  type="checkbox"
                                  checked={user.tabs.includes(tab)}
                                  disabled={tab === "User Admin"}
                                  onChange={() => toggleUserTab(user.id, tab)}
                                  className="h-4 w-4"
                                />
                                <span>{tab}</span>
                              </label>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function LoginScreen({
  authView,
  loginForm,
  loginError,
  registerForm,
  registerError,
  registerSuccess,
  users,
  onChangeLogin,
  onChangeRegister,
  onChangeView,
  onSubmitLogin,
  onSubmitRegister,
  workspaceMode,
  onEnterDemo,
  onExitDemo,
}) {
  return (
    <div className="min-h-screen bg-stone-100 p-6 text-stone-900">
      <div className="mx-auto max-w-xl rounded-[2rem] border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-8 shadow-sm shadow-stone-300/40">
        <div className="mb-6">
          <p className="text-xs font-semibold uppercase tracking-[0.18em] text-emerald-900">Secure Access</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Label Traxx Scheduler login</h1>
          <p className="mt-2 text-sm text-stone-700">
            {workspaceMode === "demo"
              ? "This is the demo workspace. It uses fake sample jobs, requests, and shipment history only."
              : "Sign in to open the scheduler, save your work, and stamp finished jobs and request history with your account."}
          </p>
        </div>
        <div className="mb-5 grid grid-cols-2 gap-2 rounded-2xl bg-stone-200/70 p-1">
          <button
            type="button"
            onClick={() => onChangeView("login")}
            className={`rounded-xl px-4 py-2 text-sm font-medium transition ${
              authView === "login" ? "bg-white text-stone-900 shadow-sm" : "text-stone-700"
            }`}
          >
            Log in
          </button>
          <button
            type="button"
            onClick={() => onChangeView("register")}
            className={`rounded-xl px-4 py-2 text-sm font-medium transition ${
              authView === "register" ? "bg-white text-stone-900 shadow-sm" : "text-stone-700"
            }`}
          >
            Register
          </button>
        </div>
        {authView === "login" ? (
          <form onSubmit={onSubmitLogin} className="grid gap-4">
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Username</div>
              <select
                value={loginForm.username}
                onChange={(event) => onChangeLogin((current) => ({ ...current, username: event.target.value }))}
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              >
                <option value="">Select a username</option>
                {users.map((user) => (
                  <option key={user.id} value={user.username}>
                    {user.username}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Password</div>
              <input
                type="password"
                value={loginForm.password}
                onChange={(event) => onChangeLogin((current) => ({ ...current, password: event.target.value }))}
                placeholder="Enter password"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            {loginError && <div className="rounded-2xl bg-rose-50 px-4 py-3 text-sm text-rose-700">{loginError}</div>}
            {registerSuccess && (
              <div className="rounded-2xl bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{registerSuccess}</div>
            )}
            <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
              Log in
            </button>
            <button
              type="button"
              onClick={workspaceMode === "demo" ? onExitDemo : onEnterDemo}
              className="rounded-2xl border border-sky-200 bg-sky-50 px-4 py-3 text-sm font-medium text-sky-900"
            >
              {workspaceMode === "demo" ? "Return to live workspace" : "Open demo workspace"}
            </button>
            {workspaceMode === "demo" && (
              <div className="rounded-2xl bg-sky-50 px-4 py-3 text-sm text-sky-900">
                Demo login: <span className="font-semibold">{DEMO_DEFAULT_USERNAME}</span> /{" "}
                <span className="font-semibold">{DEMO_DEFAULT_PASSWORD}</span>
              </div>
            )}
          </form>
        ) : (
          <form onSubmit={onSubmitRegister} className="grid gap-4">
            <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-700">
              Create a name and password. A manager will need to approve the account before you can sign in.
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Name</div>
              <input
                type="text"
                value={registerForm.username}
                onChange={(event) =>
                  onChangeRegister((current) => ({ ...current, username: event.target.value }))
                }
                placeholder="Enter your name"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Password</div>
              <input
                type="password"
                value={registerForm.password}
                onChange={(event) =>
                  onChangeRegister((current) => ({ ...current, password: event.target.value }))
                }
                placeholder="Create a password"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            {registerError && (
              <div className="rounded-2xl bg-rose-50 px-4 py-3 text-sm text-rose-700">{registerError}</div>
            )}
            <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
              Send approval request
            </button>
          </form>
        )}
        <div className="mt-4 rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-700">
          {workspaceMode === "demo" ? (
            <span>The demo site does not send or receive any live company emails.</span>
          ) : (
            <span>
              Need to make an account? Email <span className="font-semibold">sinthavong@data-mail.com</span>
            </span>
          )}
        </div>
      </div>
    </div>
  );
}

function Detail({ label, value }) {
  return (
    <div className="rounded-2xl bg-stone-100 p-3">
      <div className="text-[11px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 text-sm font-medium text-stone-800">{value}</div>
    </div>
  );
}

function Field({ label, value, onChange, placeholder }) {
  return (
    <div>
      <div className="mb-2 text-sm font-medium text-stone-800">{label}</div>
      <input
        type="text"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
      />
    </div>
  );
}

function StatRow({ label, value }) {
  return (
    <div className="rounded-2xl bg-stone-100 px-4 py-3">
      <div className="text-[11px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 text-lg font-semibold text-stone-900">{value}</div>
    </div>
  );
}

function AttachmentList({ attachments, onRemove = null }) {
  if (!attachments.length) {
    return (
      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
        No documents attached yet.
      </div>
    );
  }

  return (
    <div className="space-y-2">
      {attachments.map((attachment) => {
        const href = getAttachmentHref(attachment);
        return (
          <div key={attachment.id} className="flex flex-col gap-3 rounded-2xl border border-stone-300 bg-white p-3 md:flex-row md:items-center md:justify-between">
            <div className="min-w-0">
              {href ? (
                <a
                  href={href}
                  download={attachment.publicUrl ? undefined : attachment.name}
                  target="_blank"
                  rel="noreferrer"
                  className="block truncate text-sm font-semibold text-stone-900 underline-offset-2 hover:underline"
                >
                  {attachment.name}
                </a>
              ) : (
                <div className="block truncate text-sm font-semibold text-stone-900">{attachment.name}</div>
              )}
              <div className="mt-1 text-xs text-stone-600">
                {formatFileSize(attachment.size)} • uploaded {formatDateTime(attachment.uploadedAt)}
                {attachment.uploadedBy ? ` by ${attachment.uploadedBy}` : ""}
              </div>
            </div>
            {onRemove && (
              <button
                onClick={() => onRemove(attachment.id)}
                className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700"
              >
                Remove
              </button>
            )}
          </div>
        );
      })}
    </div>
  );
}

function RequestCard({
  request,
  onDone,
  onDelete,
  onAddAttachments,
  onRemoveAttachment,
  readOnly = false,
}) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3">
        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-sm font-semibold">
                {request.customer} - {request.jobNumber}
              </div>
              <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">
                {request.requestType || "General Request"}
              </span>
              {request.assignedToAccount && (
                <span className="rounded-full bg-sky-100 px-2 py-1 text-[11px] font-medium text-sky-900">
                  For {request.assignedToAccount}
                </span>
              )}
              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(request.status)}`}>
                {request.status}
              </span>
            </div>
            <div className="mt-1 text-xs text-stone-600">
              Requested by {request.requestorName} on {formatDateTime(request.createdAt)}
              {request.createdByAccount ? ` using ${request.createdByAccount}` : ""}
            </div>
            {request.completedAt && (
              <div className="mt-1 text-xs text-stone-600">
                Completed on {formatDateTime(request.completedAt)}
                {request.completedByAccount ? ` by ${request.completedByAccount}` : ""}
              </div>
            )}
            <div className="mt-3 whitespace-pre-wrap text-sm text-stone-800">{request.description}</div>
          </div>
          {!readOnly && (
            <div className="flex gap-2">
              <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm text-white">
                Mark done
              </button>
              <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
                Delete
              </button>
            </div>
          )}
        </div>

        <div>
          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
          <AttachmentList
            attachments={request.attachments || []}
            onRemove={readOnly ? null : onRemoveAttachment}
          />
        </div>

        {!readOnly && onAddAttachments && (
          <label className="block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
            <div className="font-medium text-stone-900">Add more files to this request</div>
            <input
              type="file"
              accept={ATTACHMENT_ACCEPT}
              multiple
              onChange={(event) => {
                const files = event.target.files;
                if (files?.length) onAddAttachments(files);
                event.target.value = "";
              }}
              className="mt-3 block w-full text-xs"
            />
          </label>
        )}
      </div>
    </div>
  );
}

function JobCard({
  job,
  state,
  onClick,
  onDoubleClick,
  onQuickAssign,
  onFinish,
  onUndoFinish,
  onUpdatePress,
  scheduledAssignments = [],
  pressOptions = [],
  weekColumns,
  finishedAt,
  finishedBy,
  canMove = true,
  selected = false,
}) {
  const isDraggable = state !== "finished" && canMove;
  const cardTone = selected
    ? "border-sky-300 bg-sky-50 shadow-sky-100/70 ring-1 ring-sky-200"
    : job.holdActive
      ? "border-rose-300 bg-rose-50 shadow-rose-100/70"
      : "border-stone-300 bg-white shadow-stone-300/20";
  const suppressClickUntilRef = useRef(0);
  const handleCardClick = () => {
    if (Date.now() < suppressClickUntilRef.current) return;
    onClick?.();
  };
  const startQueueDrag = (event) => {
    if (!isDraggable) return;
    suppressClickUntilRef.current = Date.now() + 250;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData(
      "application/json",
      makeDragPayload({ type: "queue", jobId: job.id })
    );
    event.dataTransfer.setData("text/plain", job.id);
  };

  return (
    <div
      draggable={isDraggable}
      onDragStart={startQueueDrag}
      onDragEnd={() => {
        suppressClickUntilRef.current = Date.now() + 150;
      }}
      onDoubleClick={() => onDoubleClick?.()}
      className={`rounded-2xl border p-3 shadow-sm transition-colors ${cardTone} ${isDraggable ? "cursor-grab" : ""}`}
    >
      <div className="mb-2 flex items-start justify-between gap-2">
        <div
          onClick={handleCardClick}
          onDoubleClick={onDoubleClick}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") handleCardClick();
          }}
          className="min-w-0 flex-1 cursor-pointer text-left"
        >
          <div className="text-sm font-semibold leading-tight text-stone-900">
            {job.customerName} {job.number}
          </div>
          <div className="mt-1 line-clamp-2 text-xs leading-5 text-stone-700">{job.generalDescr}</div>
        </div>
        <div className="flex flex-col items-end gap-2">
          <span className={`rounded-full border px-2 py-1 text-[11px] font-medium ${priorityTone(job.priority)}`}>
            {job.priority || "-"}
          </span>
          {job.holdActive && (
            <span className="rounded-full bg-rose-700 px-2 py-1 text-[10px] font-medium text-white">
              hold
            </span>
          )}
          {isDraggable && (
            <span className="rounded-full bg-stone-200 px-2 py-1 text-[10px] font-medium text-stone-700">
              drag
            </span>
          )}
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2 text-xs text-stone-700">
        <InfoPill label="Press" value={job.press || "-"} />
        <InfoPill label="EST" value={`${job.estPressTime.toFixed(2)}h`} />
        <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
        <InfoPill label="Ship" value={formatDate(job.shipByDate)} />
      </div>

      {onUpdatePress && pressOptions.length > 0 && (
        <div className="mt-2">
          <div className="mb-1 text-[10px] font-semibold uppercase tracking-[0.16em] text-stone-600">Recommended press</div>
          <select
            value={PRESS_ORDER.includes(job.press) ? job.press : ""}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) => {
              event.stopPropagation();
              onUpdatePress(event.target.value);
            }}
            className="w-full rounded-xl border border-stone-300 bg-stone-50 px-3 py-2 text-xs outline-none focus:border-emerald-800"
          >
            {pressOptions.map((press) => (
              <option key={press} value={press}>
                {formatPressLabel(press)}
              </option>
            ))}
          </select>
        </div>
      )}

      <div className="mt-2 flex flex-wrap items-center gap-2">
        <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(state)}`}>{state}</span>
        {job.ticketStatus && (
          <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-700">
            import: {job.ticketStatus}
          </span>
        )}
      </div>

      {scheduledAssignments.length > 0 && (
        <div className="mt-3 rounded-2xl border border-stone-200 bg-stone-50 p-3">
          <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-stone-600">Scheduled on</div>
          <div className="mt-2 flex flex-wrap gap-2">
            {scheduledAssignments.map((assignment) => (
              <span key={assignment.id} className="rounded-full bg-white px-2 py-1 text-[11px] font-medium text-stone-800">
                {formatShortDate(new Date(`${assignment.dayKey}T12:00:00`))} - {formatPressLabel(assignment.press)}
              </span>
            ))}
          </div>
        </div>
      )}

      {job.holdNote && (
        <div className="mt-2 rounded-2xl border border-rose-200 bg-white/80 p-3 text-xs text-rose-900">
          <div className="font-semibold uppercase tracking-[0.14em]">Hold note</div>
          <div className="mt-1 whitespace-pre-wrap">{job.holdNote}</div>
        </div>
      )}

      {finishedAt && <div className="mt-2 text-xs text-stone-600">Finished {formatDateTime(finishedAt)}</div>}
      {finishedBy && <div className="mt-1 text-xs text-stone-600">Finished by {finishedBy}</div>}

      {(state !== "finished" || onUndoFinish) && ((weekColumns && canMove) || onFinish || onUndoFinish) && (
        <div className="mt-3 flex flex-wrap gap-2">
          {weekColumns && canMove &&
            weekColumns.map((day) => (
              <button
                key={day.key}
                onClick={(event) => {
                  event.stopPropagation();
                  onQuickAssign?.(day.key);
                }}
                className="rounded-xl border border-stone-300 bg-stone-50 px-2 py-1 text-[11px] text-stone-800 hover:bg-stone-100"
              >
                {day.label.slice(0, 3)}
              </button>
            ))}
          {onFinish && (
            <button
              onClick={(event) => {
                event.stopPropagation();
                onFinish();
              }}
              className="rounded-xl bg-emerald-900 px-2 py-1 text-[11px] text-white"
            >
              Finish
            </button>
          )}
          {onUndoFinish && (
            <button
              onClick={(event) => {
                event.stopPropagation();
                onUndoFinish();
              }}
              className="rounded-xl border border-emerald-300 bg-white px-2 py-1 text-[11px] text-emerald-950"
            >
              Unmark done
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function PaperPullCard({ request, onDone, onDelete, readOnly = false }) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
        <div className="space-y-2">
          <div className="flex flex-wrap items-center gap-2">
            <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">{request.target}</span>
            <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${request.status === "done" ? statusTone("done") : statusTone("open")}`}>
              {request.status}
            </span>
          </div>
          <div className="whitespace-pre-wrap text-sm text-stone-800">{request.details}</div>
          <div className="text-xs text-stone-600">
            Requested by {request.createdBy || "-"} on {formatDateTime(request.createdAt)}
          </div>
          {request.completedAt && (
            <div className="text-xs text-stone-600">
              Completed by {request.completedBy || "-"} on {formatDateTime(request.completedAt)}
            </div>
          )}
        </div>
        {!readOnly && (
          <div className="flex gap-2">
            <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white">
              Done
            </button>
            <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
              Delete
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function PressQueueRow({
  job,
  state,
  selected = false,
  onSelect,
  onOpenDetails,
  onFinish,
  canMove = false,
  scheduledAssignments = [],
  pressOptions = [],
  onUpdatePress,
  weekColumns = [],
  onQuickAssign,
  onPickUp,
}) {
  const suppressClickUntilRef = useRef(0);
  const rowTone = selected
    ? "bg-sky-100"
    : job?.holdActive
      ? "bg-rose-50"
      : "bg-white even:bg-stone-50";
  const scheduledText = scheduledAssignments.length
    ? scheduledAssignments
        .map((assignment) => `${formatShortDate(new Date(`${assignment.dayKey}T12:00:00`))} ${formatPressLabel(assignment.press)}`)
        .join(", ")
    : "-";

  const startQueueDrag = (event) => {
    if (!canMove) return;
    suppressClickUntilRef.current = Date.now() + 250;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("application/json", makeDragPayload({ type: "queue", jobId: job.id }));
    event.dataTransfer.setData("text/plain", job.id);
  };

  const handleSelect = () => {
    if (Date.now() < suppressClickUntilRef.current) return;
    onSelect?.();
  };

  return (
    <tr
      onClick={handleSelect}
      onDoubleClick={onOpenDetails}
      className={`cursor-pointer border-b border-stone-200 align-top ${rowTone} hover:bg-sky-50`}
    >
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2 font-medium">{job.number}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.customerName}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.priority || "-"}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.custPoNum || "-"}</td>
      <td className="max-w-[280px] truncate border-r border-stone-200 px-3 py-2" title={job.generalDescr}>
        {job.generalDescr}
      </td>
      <td className="whitespace-nowrap border-r border-stone-200 px-2 py-2">
        {onUpdatePress && pressOptions.length > 0 ? (
          <select
            value={PRESS_ORDER.includes(job.press) ? job.press : ""}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) => {
              event.stopPropagation();
              onUpdatePress(event.target.value);
            }}
            className="w-[90px] border border-stone-300 bg-white px-2 py-1 text-[12px] outline-none focus:border-emerald-800"
          >
            {pressOptions.map((press) => (
              <option key={press} value={press}>
                {press}
              </option>
            ))}
          </select>
        ) : (
          job.press || "-"
        )}
      </td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{formatDate(job.shipByDate)}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{formatDate(job.dueOnSiteDate)}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.ticQuantity ? job.ticQuantity.toLocaleString() : "-"}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">
        <span className={selected ? "font-semibold text-sky-800" : ""}>{state}</span>
      </td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.stockDisplay || "-"}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.estPressTime ? job.estPressTime.toFixed(2) : "-"}</td>
      <td className="whitespace-nowrap border-r border-stone-200 px-3 py-2">{job.mainTool || job.toolNo2 || "-"}</td>
      <td className="max-w-[260px] truncate border-r border-stone-200 px-3 py-2" title={scheduledText}>
        {scheduledText}
      </td>
      <td className="border-r border-stone-200 px-2 py-2">
        <div className="flex flex-wrap gap-1">
          {weekColumns.map((day) => (
            <button
              key={day.key}
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                onQuickAssign?.(day.key);
              }}
              className="border border-stone-300 bg-white px-1.5 py-1 text-[10px] text-stone-800"
            >
              {day.label.slice(0, 3)}
            </button>
          ))}
        </div>
      </td>
      <td className="whitespace-nowrap px-3 py-2">
        <div className="flex items-center gap-2">
          {onFinish && (
            <button
              onClick={(event) => {
                event.stopPropagation();
                onFinish();
              }}
              className="bg-emerald-900 px-2 py-1 text-[11px] text-white"
            >
              Finish
            </button>
          )}
          {canMove && (
            <>
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onPickUp?.();
                }}
                className="border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-700"
              >
                Pick up
              </button>
              <span
                draggable
                onDragStart={startQueueDrag}
                onDragEnd={() => {
                  suppressClickUntilRef.current = Date.now() + 150;
                }}
                onClick={(event) => event.stopPropagation()}
                className="border border-stone-300 bg-stone-100 px-2 py-1 text-[11px] text-stone-700"
              >
                drag
              </span>
            </>
          )}
        </div>
      </td>
    </tr>
  );
}

function SuppliesRequestCard({
  request,
  onDone,
  onDelete,
  onAddAttachments,
  onRemoveAttachment,
  readOnly = false,
}) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3">
        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-sm font-semibold">Supplies request</div>
              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(request.status)}`}>
                {request.status}
              </span>
            </div>
            <div className="mt-1 text-xs text-stone-600">
              Requested by {request.createdBy || "-"} on {formatDateTime(request.createdAt)}
            </div>
            {request.completedAt && (
              <div className="mt-1 text-xs text-stone-600">
                Completed by {request.completedBy || "-"} on {formatDateTime(request.completedAt)}
              </div>
            )}
            <div className="mt-3 whitespace-pre-wrap text-sm text-stone-800">{request.details}</div>
          </div>
          {!readOnly && (
            <div className="flex gap-2">
              <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white">
                Done
              </button>
              <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
                Delete
              </button>
            </div>
          )}
        </div>

        <div>
          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
          <AttachmentList attachments={request.attachments || []} onRemove={readOnly ? null : onRemoveAttachment} />
        </div>

        {!readOnly && onAddAttachments && (
          <label className="block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
            <div className="font-medium text-stone-900">Add more files to this supplies request</div>
            <input
              type="file"
              accept={ATTACHMENT_ACCEPT}
              multiple
              onChange={(event) => {
                const files = event.target.files;
                if (files?.length) onAddAttachments(files);
                event.target.value = "";
              }}
              className="mt-3 block w-full text-xs"
            />
          </label>
        )}
      </div>
    </div>
  );
}

function CompactScheduleCard({
  job,
  assignment,
  finishMeta,
  density = "compact",
  selected = false,
  onSelect,
  canMoveUp = false,
  canMoveDown = false,
  onMoveUp,
  onMoveDown,
  onDropBefore,
  onUnschedule,
  onFinish,
  onUndoFinish,
  onDuplicate,
  draggable = false,
  onPickUp,
}) {
  const isManual = assignment.kind === "manual";
  const state = isManual ? "note" : assignment.status === "finished" ? "done" : assignment.status;
  const title = isManual ? assignment.manualTitle || "Manual block" : `${job.customerName} ${job.number}`;
  const subtitle = isManual ? "Manual schedule block" : job.generalDescr;
  const isCompact = density === "compact";
  const showSubtitle = density === "detailed";
  const showStateBadge = isManual || assignment.status === "finished";
  const showStats = !isManual && density === "detailed";
  const showFinishedStamp = density === "detailed" && !isManual;
  const isCardDraggable = draggable;
  const showReorderControls = canMoveUp || canMoveDown || onMoveUp || onMoveDown;
  const cardTone = selected
    ? "border-sky-300 bg-sky-50 ring-1 ring-sky-200"
    : job?.holdActive
      ? "border-rose-300 bg-rose-50"
      : "border-stone-200 bg-white";
  const suppressClickUntilRef = useRef(0);
  const handleSelect = () => {
    if (Date.now() < suppressClickUntilRef.current) return;
    onSelect?.();
  };
  const startScheduledDrag = (event) => {
    if (!isCardDraggable) return;
    suppressClickUntilRef.current = Date.now() + 250;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData(
      "application/json",
      makeDragPayload({ type: "scheduled", assignmentId: assignment.id, jobId: job?.id || "" })
    );
  };

  return (
    <div
      draggable={isCardDraggable}
      onDragStart={startScheduledDrag}
      onDragEnd={() => {
        suppressClickUntilRef.current = Date.now() + 150;
      }}
      onDragOver={
        onDropBefore
          ? (event) => {
              event.preventDefault();
              event.stopPropagation();
            }
          : undefined
      }
      onDrop={onDropBefore}
      className={`border-b p-2 transition-colors ${cardTone} ${isCardDraggable ? "cursor-grab" : ""}`}
    >
      <div className="flex items-start justify-between gap-2">
        <div
          onClick={handleSelect}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") handleSelect();
          }}
          className={`min-w-0 flex-1 ${onSelect ? "cursor-pointer" : ""}`}
        >
          <div className="truncate text-xs font-semibold text-stone-900">
            {title}
          </div>
          {showSubtitle && <div className="mt-1 line-clamp-2 text-[11px] text-stone-700">{subtitle}</div>}
        </div>
        <div className="flex items-start gap-2">
          {showReorderControls && (
            <div className="flex items-center gap-1">
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onMoveUp?.();
                }}
                disabled={!canMoveUp}
                className={`border px-2 py-1 text-[10px] font-medium ${canMoveUp ? "border-stone-300 bg-white text-stone-800" : "border-stone-200 bg-stone-100 text-stone-400"}`}
                aria-label="Move schedule item up">
                &uarr;
              </button>
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onMoveDown?.();
                }}
                disabled={!canMoveDown}
                className={`border px-2 py-1 text-[10px] font-medium ${canMoveDown ? "border-stone-300 bg-white text-stone-800" : "border-stone-200 bg-stone-100 text-stone-400"}`}
                aria-label="Move schedule item down">
                &darr;
              </button>
            </div>
          )}
          {job?.holdActive && (
            <span className="bg-rose-700 px-2 py-1 text-[10px] font-medium text-white">hold</span>
          )}
          {showStateBadge && <span className={`px-2 py-1 text-[10px] font-medium ${statusTone(state)}`}>{state}</span>}
          {isCardDraggable && <span className="border border-stone-300 bg-stone-100 px-2 py-1 text-[10px] font-medium text-stone-700">drag</span>}
        </div>
      </div>
      {!isManual && isCompact && (
        <div className="mt-2 text-[11px] text-stone-700">
          Ship {formatDate(job.shipByDate)}
        </div>
      )}
      {showStats && (
        <div className="mt-2 grid grid-cols-2 gap-1 text-[11px] text-stone-700 md:grid-cols-3">
          <InfoPill label="Est" value={`${job.estPressTime.toFixed(2)}h`} />
          <InfoPill label="Footage" value={job.estFootage.toLocaleString()} />
          <InfoPill label="Stock" value={job.stockDisplay || "-"} />
          <InfoPill label="Ship" value={formatDate(job.shipByDate)} />
          <InfoPill label="Priority" value={job.priority || "-"} />
          <InfoPill label="Die" value={job.mainTool || job.toolNo2 || "-"} />
        </div>
      )}
      {job?.holdNote && (
        <div className="mt-2 border border-rose-200 bg-rose-50 p-2 text-[11px] text-rose-900">
          <div className="font-semibold uppercase tracking-[0.14em]">Hold note</div>
          <div className="mt-1 whitespace-pre-wrap">{job.holdNote}</div>
        </div>
      )}
      {finishMeta?.finishedAt && showFinishedStamp && (
        <div className="mt-2 text-[11px] text-stone-600">
          {formatDateTime(finishMeta.finishedAt)} {finishMeta.finishedBy ? `- ${finishMeta.finishedBy}` : ""}
        </div>
      )}
      {(onUnschedule || onFinish || onUndoFinish || onDuplicate) && (
        <div className={`${density === "compact" ? "mt-1" : "mt-2"} flex flex-wrap gap-2`}>
          {onPickUp && (
            <button onClick={onPickUp} className="border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800">
              Pick up
            </button>
          )}
          {onUnschedule && (
            <button onClick={onUnschedule} className="border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800">
              Remove
            </button>
          )}
          {onDuplicate && (
            <button onClick={onDuplicate} className="border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800">
              Duplicate
            </button>
          )}
          {onFinish && (
            <button onClick={onFinish} className="bg-emerald-900 px-2 py-1 text-[11px] text-white">
              Finish
            </button>
          )}
          {onUndoFinish && (
            <button onClick={onUndoFinish} className="border border-emerald-300 bg-white px-2 py-1 text-[11px] text-emerald-950">
              Unmark done
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function InfoPill({ label, value }) {
  return (
    <div className="border border-stone-200 bg-stone-50 px-2 py-2">
      <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 font-medium text-stone-800">{value}</div>
    </div>
  );
}

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error) {
    console.error("Label Traxx Scheduler crashed.", error);
  }

  resetBrowserData = () => {
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(SESSION_STORAGE_KEY);
    window.location.reload();
  };

  render() {
    if (!this.state.error) return this.props.children;

    return (
      <div className="min-h-screen bg-stone-100 p-6 text-stone-900">
        <div className="mx-auto max-w-2xl rounded-[2rem] border border-rose-200 bg-white p-8 shadow-sm shadow-stone-300/40">
          <p className="text-xs font-semibold uppercase tracking-[0.18em] text-rose-700">App error</p>
          <h1 className="mt-2 text-2xl font-semibold tracking-tight">The scheduler hit a browser-side error.</h1>
          <p className="mt-3 text-sm text-stone-700">
            Use the reset button below to clear saved browser data for this device and reload the app.
          </p>
          <div className="mt-4 rounded-2xl bg-stone-100 p-4 text-sm text-stone-800">
            {safeText(this.state.error?.message) || "Unknown error"}
          </div>
          <div className="mt-5 flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => window.location.reload()}
              className="rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800"
            >
              Reload app
            </button>
            <button
              type="button"
              onClick={this.resetBrowserData}
              className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white"
            >
              Reset saved browser data
            </button>
          </div>
        </div>
      </div>
    );
  }
}

export default function App() {
  return (
    <AppErrorBoundary>
      <SchedulerApp />
    </AppErrorBoundary>
  );
}

